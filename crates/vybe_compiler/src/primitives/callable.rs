//! Shared callable invocation emitters.
//!
//! This is the common boundary for values that are *called*: lambdas,
//! function references, method references and delegate values. Object
//! interception belongs in the proxy primitive; proxy traps can reuse this
//! module because each trap is itself a callable.

use vybe_runtime::Chunk;

use crate::primitives::class_slots;
use super::Compiler;

/// Callable invocation flavor. The argument count excludes the callable value
/// already on the stack, matching `CALL_REF` and ordinary language syntax.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InvokeKind {
    Direct,
    MulticastDelegate,
}

/// Invoke the callable value currently below `arg_count` arguments on the
/// stack. Stack before: `[callable, arg0, ...]`; after: `[result]`.
pub fn emit_invoke(
    chunks: &mut [Chunk],
    current: usize,
    arg_count: u8,
    kind: InvokeKind,
    line: u32,
) {
    match kind {
        InvokeKind::Direct => emit_direct_invoke(chunks, current, arg_count, line),
        InvokeKind::MulticastDelegate => {
            emit_multicast_delegate_invoke(chunks, current, arg_count, line)
        }
    }
}

/// Direct lambda/function-reference call. The callable and its arguments must
/// already be on the stack.
pub fn emit_direct_invoke(chunks: &mut [Chunk], current: usize, arg_count: u8, line: u32) {
    emit_direct_invoke_chunk(&mut chunks[current], arg_count, line);
}

/// Push the RECEIVER SLOT of a callback invocation, and report how much it
/// widens the argument count.
///
/// ⛔ CALL IT AFTER THE CALLEE IS ON THE STACK AND BEFORE ITS ARGUMENTS.
/// §10.2.1 `[[Call]](thisArgument, argumentsList)` puts the receiver at
/// argument 0, so it cannot be added by the invoke helper — by then the
/// arguments are already stacked above it.
///
/// ⛔ WHY THIS EXISTS AT ALL: a callback callee compiled under
/// `UniversalParameter` DECLARES a receiver parameter, so a call site that
/// pushes only the real arguments leaves every one of them arriving a place
/// early. Measured on dart the moment its directive flipped:
/// `[10,20,30].map((e) => e)` answered `0,1,2` — the element landed in the
/// receiver slot and `e` took the index. Under the ambient ABI this emits
/// nothing and returns 0, so the shape is unchanged.
///
/// Returns the number of EXTRA arguments (0 or 1) to add to the invoke's
/// `arg_count`.
pub fn emit_callback_receiver(
    chunk: &mut Chunk,
    abi: vybe_runtime::chunk::ReceiverAbi,
    line: u32,
) -> u8 {
    if abi != vybe_runtime::chunk::ReceiverAbi::Parameter {
        return 0;
    }
    // A method reference is a callable value that carries its receiver in the
    // shared class slot. Plain functions/lambdas have no such slot, and the
    // read naturally yields `undefined`, matching §10.2.1.1.
    chunk.emit_dup(line);
    let receiver_slot = class_slots::resolve(
        &class_slots::ClassSlot::internal("__vybe_method_receiver"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &receiver_slot,
        class_slots::Dest::Stack,
        line,
    );
    1
}

/// Push a callee held in a local together with its receiver; returns the extra
/// argument count (0 or 1).
///
/// ⛔ USE THIS RATHER THAN A BARE `local.get` OF THE CALLEE. §10.2.1 puts the
/// receiver at argument 0, so once the real arguments are stacked it is too
/// late to add one. Returns 0 and emits nothing extra where the region
/// declares no receiver.
pub fn push_callback_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    callee_slot: u16,
    line: u32,
) -> u8 {
    let abi = crate::primitives::class_context::module_receiver_abi(chunks);
    chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, callee_slot, line);
    emit_callback_receiver(&mut chunks[current], abi, line)
}

/// Invoke the callback in `fn_slot` on the value in `arg_slot`, leaving its
/// result on the stack.
///
/// ⛔ THE ONE PLACE A CALLBACK LOOP PLACES ITS RECEIVER. §10.2.1 puts it at
/// argument 0, which is not expressible once the argument is already stacked,
/// so an emitter that writes these four instructions itself cannot be given one
/// later. Call this rather than emitting the invoke.
pub fn emit_callback_on(chunks: &mut [Chunk], current: usize, fn_slot: u16, arg_slot: u16, line: u32) {
    let recv = push_callback_from_slot(chunks, current, fn_slot, line);
    chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, arg_slot, line);
    emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
}

/// [`emit_callback_on`] for a two-argument callback — a selector over
/// `(element, index)`, a result selector over `(outer, inner)`, a fold step
/// over `(accumulator, element)`.
pub fn emit_callback_on2(
    chunks: &mut [Chunk],
    current: usize,
    fn_slot: u16,
    a_slot: u16,
    b_slot: u16,
    line: u32,
) {
    let recv = push_callback_from_slot(chunks, current, fn_slot, line);
    chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, a_slot, line);
    chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, b_slot, line);
    emit_direct_invoke_chunk(&mut chunks[current], 2 + recv, line);
}

/// Direct lambda/function-reference call for helpers that are already building
/// a function chunk rather than writing into the active chunk vector.
pub fn emit_direct_invoke_chunk(chunk: &mut Chunk, arg_count: u8, line: u32) {
    crate::primitives::functions::emit_call(chunk, arg_count, line);
}

/// Invoke a callable whose arguments are ALREADY stacked above it, inserting the
/// receiver at argument 0 where the region declares one.
///
/// ⛔ THE RECEIVER IS ARGUMENT 0, SO IT CANNOT BE APPENDED. A helper that has
/// already pushed `[callee, arg0..argN]` has no room left underneath, so the
/// arguments are spilled to scratch, `undefined` is pushed (§10.2.1.1
/// OrdinaryCallBindThis — a call with no receiver of its own binds `undefined`)
/// and the arguments are restacked above it. Under `Ambient` this is the plain
/// invoke and emits no extra instruction.
pub fn emit_stacked_invoke(chunks: &mut [Chunk], current: usize, arg_count: u8, line: u32) {
    let abi = crate::primitives::class_context::module_receiver_abi(chunks);
    emit_stacked_invoke_chunk(&mut chunks[current], abi, arg_count, line);
}

/// [`emit_stacked_invoke`] for a helper holding a single chunk.
///
/// ⛔ THE CONVENTION IS GIVEN, NEVER LOOKED UP BY POSITION. A caller holding one
/// chunk cannot ask the module for its ABI — `chunks.first()` would be a
/// function chunk and would answer the default for every module — so it states
/// what it is building for.
pub fn emit_stacked_invoke_chunk(
    chunk: &mut Chunk,
    abi: vybe_runtime::chunk::ReceiverAbi,
    arg_count: u8,
    line: u32,
) {
    if abi != vybe_runtime::chunk::ReceiverAbi::Parameter {
        emit_direct_invoke_chunk(chunk, arg_count, line);
        return;
    }
    let slots: Vec<u16> = (0..arg_count).map(|_| chunk.alloc_scratch(1)).collect();
    for slot in slots.iter().rev() {
        chunk.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, *slot, line);
    }
    crate::primitives::expressions::emit_undefined(chunk, line);
    for slot in slots.iter() {
        chunk.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, *slot, line);
    }
    emit_direct_invoke_chunk(chunk, arg_count.saturating_add(1), line);
}

/// .NET-style multicast delegate call. The public callable contract keeps
/// `arg_count` excluding the delegate; the underlying delegate emitter uses an
/// older stack-width convention, so adapt it here once.
pub fn emit_multicast_delegate_invoke(
    chunks: &mut [Chunk],
    current: usize,
    arg_count: u8,
    line: u32,
) {
    crate::primitives::delegates::emit_invoke(chunks, current, arg_count.saturating_add(1), line);
}

impl Compiler {
    pub(crate) fn emit_callable_invoke(&mut self, arg_count: u8, kind: InvokeKind) {
        emit_invoke(&mut self.chunks, self.current, arg_count, kind, self.line);
    }

    pub(crate) fn emit_direct_callable_invoke(&mut self, arg_count: u8) {
        self.emit_callable_invoke(arg_count, InvokeKind::Direct);
    }

    pub(crate) fn emit_multicast_delegate_invoke(&mut self, arg_count: u8) {
        self.emit_callable_invoke(arg_count, InvokeKind::MulticastDelegate);
    }
}
