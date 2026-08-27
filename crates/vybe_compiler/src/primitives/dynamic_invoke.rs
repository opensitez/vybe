//! Dynamic invocation on a dynamically typed receiver.
//!
//! `receiver.foo(arg)` where the receiver's type is unknown at the call site,
//! and `foo` may not exist on it at all. The Call Tags proposal names this
//! application directly (`proposals/call-tags/…/Overview.md`, "Dynamic
//! Invocation"):
//!
//! > They associated a call tag […] with each name-arity pair, whose signature
//! > accepted the appropriate number of generic objects and returned a generic
//! > object. […] the fall-back handler for these call tags would search through
//! > the additional-methods dictionary of the object to find a corresponding
//! > entry and call it (if there were any).
//!
//! PHP `__call`, Ruby `method_missing` and Dart `noSuchMethod` ARE that
//! additional-methods dictionary. All four frontends that have the feature bind
//! `ProtocolSlot::CallMissing`, and until this module nothing read it — so each
//! walker synthesised its own `typeof obj.__call === "function" ? … : …`
//! rewrite, per call site, on parse pairs. This is the reader.
//!
//! ## Why the dispatch is the callee's decision
//!
//! The call site pushes `[receiver arg* funcref]` and issues one
//! `call_with_tag`. Three outcomes, all resolved by the VM:
//!
//! 1. the funcref handles the tag → it is called directly, no probe, no branch;
//! 2. it does not → the tag's fall-back handler runs, receiving the arguments
//!    unchanged plus *the funcref that refused*;
//! 3. the method did not resolve at all → the call site pushes [`SENTINEL`],
//!    which handles no tag, so (2) runs with the sentinel as the funcref.
//!
//! Case (2) carrying the funcref is what makes stamping an OPTIMISATION rather
//! than a correctness requirement: a method that exists but was not stamped for
//! this arity (defaults, variadics) arrives at the handler as a live funcref and
//! is simply called. Only the sentinel means "no such method", and only then is
//! the `CallMissing` slot consulted. A partial stamp costs speed, never an
//! answer — which matters, because "route a method that exists to `__call`" is
//! a silently wrong result.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use super::*;

/// Funcref pushed when the method did not resolve on the receiver. Declares no
/// call tags, so every tag misses it and the fall-back handler runs.
pub const SENTINEL: &str = "vybe:invoke-no-method";

/// The tag naming a (method, arity) pair. Both halves matter: the name selects
/// the method, the arity fixes the signature the tag declares.
pub fn invoke_tag_name(method: &str, argc: usize) -> String {
    format!("vybe:invoke/{method}/{argc}")
}

/// The tag's fall-back handler. One per (method, arity), because the handler
/// has to name the method when it reaches `CallMissing` and a fall-back sees
/// only values — not which tag sent it.
fn miss_handler_name(method: &str, argc: usize) -> String {
    format!("vybe:invoke-miss/{method}/{argc}")
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn is_undefined(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(idx, 1, line);
}

/// `[…] -> [ti* funcref] -> [to*]`: the handler's parameters are the tag's
/// parameters plus the refusing funcref, exactly as the Overview specifies.
///
/// Slot layout: `0` receiver, `1..=argc` arguments, `argc+1` funcref.
fn build_miss_handler(
    chunks: &mut Vec<Chunk>,
    method: &str,
    argc: usize,
    error_type: &str,
    line: u32,
) -> usize {
    let idx = chunks.len();
    let arity = (argc + 2) as u8;
    let mut h = common::functions::create_function_chunk(&miss_handler_name(method, argc), arity);
    h.alloc_scratch(arity as u16);

    let recv = 0u16;
    let funcref = (argc + 1) as u16;

    // The funcref REFUSED the tag but still exists — the method is there and
    // merely was not stamped for this arity. Call it; a stamp is a fast path,
    // not the definition of "has this method".
    lget(&mut h, funcref, line);
    is_undefined(&mut h, line);
    h.emit_op(Op::I32_EQZ, line);
    lget(&mut h, funcref, line);
    h.emit_op(Op::REF_IS_NULL, line);
    h.emit_op(Op::I32_EQZ, line);
    h.emit_op(Op::I32_AND, line);
    h.emit_if(line);
    {
        lget(&mut h, funcref, line);
        lget(&mut h, recv, line);
        for i in 0..argc {
            lget(&mut h, (i + 1) as u16, line);
        }
        common::callable::emit_direct_invoke_chunk(&mut h, (argc + 1) as u8, line);
        h.emit_op(Op::RETURN, line);
    }
    h.emit_end(line);

    // No method. Consult the receiver's `CallMissing` slot — the
    // additional-methods dictionary the Overview describes.
    let slot_key = vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::CallMissing);
    let handler = h.alloc_scratch(1);
    let get = h.add_import("ecma:reflect", "get");
    lget(&mut h, recv, line);
    h.emit_string_const(&slot_key, line);
    h.emit_call(get, 2, line);
    h.emit_op_u16(Op::LOCAL_SET, handler, line);

    lget(&mut h, handler, line);
    is_undefined(&mut h, line);
    lget(&mut h, handler, line);
    h.emit_op(Op::REF_IS_NULL, line);
    h.emit_op(Op::I32_OR, line);
    h.emit_if(line);
    {
        // The class binds no miss hook, so this really is an undefined method.
        // A catchable language error, not a VM trap.
        h.emit_struct_new(0, 0, line);
        h.emit_dup(line);
        h.emit_string_const("Call to undefined method ", line);
        h.emit_string_const(method, line);
        h.emit_string_const("()", line);
        common::strings::emit_concat(&mut h, 3, line);
        common::errors::emit_exception_new_finalize(&mut h, error_type, line);
        common::errors::emit_throw(&mut h, line);
    }
    h.emit_end(line);

    // `__call($name, $args)` — receiver-first, as every slot-bound method is
    // compiled, with the arguments collected into the language's array.
    lget(&mut h, handler, line);
    lget(&mut h, recv, line);
    h.emit_string_const(method, line);
    for i in 0..argc {
        lget(&mut h, (i + 1) as u16, line);
    }
    h.emit_array_new_fixed(0, argc as u16, line);
    common::callable::emit_direct_invoke_chunk(&mut h, 3, line);
    h.emit_op(Op::RETURN, line);

    chunks.push(h);
    idx
}

/// The no-such-method funcref. Handles no call tag, which is the whole point.
fn build_sentinel(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let idx = chunks.len();
    let mut h = common::functions::create_function_chunk(SENTINEL, 0);
    common::functions::emit_function_epilogue(&mut h, line);
    chunks.push(h);
    idx
}

fn find_chunk(chunks: &[Chunk], name: &str) -> Option<usize> {
    chunks.iter().position(|c| c.name == name)
}

impl Compiler {
    /// Chunk index of the shared no-such-method sentinel, building it once.
    pub(crate) fn invoke_sentinel_chunk(&mut self) -> usize {
        match find_chunk(&self.chunks, SENTINEL) {
            Some(i) => i,
            None => {
                let line = self.line;
                build_sentinel(&mut self.chunks, line)
            }
        }
    }

    /// Declare `vybe:invoke/<method>/<argc>` and its fall-back handler, once per
    /// pair, returning the tag's name.
    ///
    /// The declaration is recorded on the chunk being compiled; load-time
    /// resolution interns tags by NAME across every chunk, so a tag declared
    /// here and a `call_with_tag` naming it anywhere else meet at one entity.
    pub(crate) fn ensure_invoke_tag(&mut self, method: &str, argc: usize) -> String {
        let tag = invoke_tag_name(method, argc);
        let handler = miss_handler_name(method, argc);
        if find_chunk(&self.chunks, &handler).is_none() {
            let line = self.line;
            // The language's "you called a method you cannot call" error
            // class — PHP `Error`, JS `TypeError`. Deliberately the SAME
            // profile row the null-receiver member call throws, rather than a
            // new one: both are the same language-level failure, and every
            // language that has one has the other.
            let err = self.profile.member_call_on_null_error.clone();
            build_miss_handler(&mut self.chunks, method, argc, &err, line);
        }
        // params = receiver + arguments; one result.
        self.chunks[self.current].declare_call_tag(
            tag.clone(),
            (argc + 1) as u8,
            1,
            Some(handler),
            false,
        );
        tag
    }

    /// Emit the dispatch. Stack on entry: nothing; on exit: the call's result.
    ///
    /// `callee_slot` holds the resolved method, or null/undefined when it did
    /// not resolve — in which case the sentinel is pushed instead, so the
    /// unhandled-tag path runs rather than `call_with_tag` erroring on a
    /// non-function.
    pub(crate) fn emit_dynamic_invoke_with_tag(
        &mut self,
        obj_slot: u16,
        callee_slot: u16,
        method: &str,
        arg_slots: &[u16],
    ) {
        let tag = self.ensure_invoke_tag(method, arg_slots.len());
        let sentinel = self.invoke_sentinel_chunk();
        let line = self.line;

        // `[ti* funcref]` — the proposal's operand order: arguments first,
        // funcref on top, receiver as parameter 0 so the fall-back handler has
        // a `$this` to reach the miss hook on.
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        for slot in arg_slots {
            self.emit_u16(Op::LOCAL_GET, *slot);
        }

        self.emit_u16(Op::LOCAL_GET, callee_slot);
        self.emit(Op::REF_IS_NULL);
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        {
            let idx = self.chunk().add_import("wasm:js-undefined", "test");
            self.chunk().emit_call(idx, 1, line);
        }
        self.emit(Op::I32_OR);
        self.chunk().emit_if_value(line);
        self.chunk().emit_op_u16(Op::REF_FUNC, sentinel as u16, line);
        self.chunk().emit(0, line);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        self.chunk().emit_end(line);

        let tag_idx = self.str_const(&tag);
        let chunk = &mut self.chunks[self.current];
        chunk.emit_op_u16(Op::CALL_WITH_TAG, tag_idx, line);
        chunk.emit(arg_slots.len() as u8 + 1, line);
    }
}
