//! Shared callable invocation emitters.
//!
//! This is the common boundary for values that are *called*: lambdas,
//! function references, method references and delegate values. Object
//! interception belongs in the proxy primitive; proxy traps can reuse this
//! module because each trap is itself a callable.

use vybe_runtime::Chunk;

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
    // §10.2.1.1 OrdinaryCallBindThis: a call with no receiver of its own binds
    // `undefined` — "absent" and "undefined" are not the same slot.
    crate::primitives::expressions::emit_undefined(chunk, line);
    1
}

/// Direct lambda/function-reference call for helpers that are already building
/// a function chunk rather than writing into the active chunk vector.
pub fn emit_direct_invoke_chunk(chunk: &mut Chunk, arg_count: u8, line: u32) {
    crate::primitives::functions::emit_call(chunk, arg_count, line);
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
