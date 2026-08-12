//! Dynamic method invocation — the polymorphic shim for `receiver.method(args)`
//! when the receiver's type is unknown at compile time.
//!
//! # Why this exists
//!
//! JS `x.slice(0, 5)` is semantically polymorphic: strings look up
//! `slice` on `String.prototype`, arrays on `Array.prototype`, user
//! objects on their own prototype chain. The wasm js-builtins proposals
//! mirror that with **separate import modules** — `wasm:js-string.*` vs
//! `ecma:array.*` — which is correct but forces either compile-time
//! type inference or runtime dispatch for every `.slice()` call site.
//!
//! This module provides the runtime-dispatch path. The emitted bytecode
//! routes every dynamic method call through a single import:
//!
//! ```text
//! ecma:value.invokeMethod(receiver, method_name, ...args) -> value
//! ```
//!
//! On v8 via the js-builtins bridge, the glue implementation resolves
//! `receiver[method_name](...args)` — identical to the native JS lookup
//! (prototype-chain + method-missing semantics). On Vybe's own VM the
//! host handler does the same walk.
//!
//! # When to use
//!
//! Use `emit_invoke_method` for method calls where the receiver's type
//! isn't statically known. For statically-typed receivers (VB `As String`,
//! C# `List<int>`, Pascal typed parameters), emit the typed import
//! directly via `collections::emit_*` or `strings::emit_*` — those are
//! always cheaper than a polymorphic dispatch.
//!
//! # Stack contract
//!
//! Before: `[receiver, arg1, arg2, ..., argN]`
//! After : `[result]`
//!
//! The helper splices the method-name constant between `receiver` and
//! the args using temp local slots (no `SWAP` / `INSERT` opcodes).

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// Emit a polymorphic `receiver.method(args)` invocation.
///
/// Stack before:  `[receiver, arg1, arg2, ..., argN]` (argc args)
/// Stack after :  `[result]`
///
/// The emitted sequence calls `ecma:value.invokeMethod(receiver,
/// method_name, ...args)`. Argc is bounded to 253 (plus receiver +
/// name = 255, fitting in a u8).
///
/// Scratch slots are allocated starting at `chunk.local_count` —
/// **safe by construction** because `Compiler::define_local` keeps
/// `chunk.local_count >= scope.next_slot` at all times. Helpers that
/// take `&mut [Chunk]` and trust `chunk.local_count` for scratch base
/// rely on this invariant.
pub fn emit_invoke_method(
    chunks: &mut [Chunk],
    current: usize,
    method_name: &str,
    argc: u8,
    line: u32,
) {
    // Always go through the receiver-stash path so we can also bind
    // `__js_this` before the host call. The host's `invokeMethod` →
    // `dispatch` → `ctx.invoke` chain runs the user method body with
    // whatever `__js_this` is currently set to (the bridge JS-compiled
    // class methods read `this` from). Setting `__js_this = receiver`
    // here lets `obj.method()` reach a body that does `this.x` even
    // when the method is bound on the instance and dispatched
    // dynamically.
    let c = &mut chunks[current];
    let temp_base = c.alloc_scratch(argc as u16 + 2);
    let receiver_slot = temp_base;
    let prev_this_slot = temp_base + 1;
    let arg_base = temp_base + 2;

    // Pop args into temps (LIFO: last arg lands in highest temp slot).
    for i in (0..argc).rev() {
        let slot = arg_base + i as u16;
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    // Stash receiver.
    c.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);

    // Save current __js_this so we can restore after the call. Host
    // functions don't manage this global — every JS method call site
    // (here and in primitives/calls.rs) is responsible for save/restore.
    crate::primitives::globals::emit_read(c, "__js_this", line);
    c.emit_op_u16(Op::LOCAL_SET, prev_this_slot, line);

    // Set __js_this = receiver so JS-compiled method bodies see the
    // right `this` when dispatch eventually drives `ctx.invoke`.
    c.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    crate::primitives::globals::emit_write(c, "__js_this", line);

    // Rebuild call stack: receiver, name, args...
    c.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    c.emit_string_const(method_name, line);
    for i in 0..argc {
        let slot = arg_base + i as u16;
        c.emit_op_u16(Op::LOCAL_GET, slot, line);
    }

    let idx = chunks[current].add_import("ecma:value", "invokeMethod");
    let c = &mut chunks[current];
    c.emit_call(idx, argc + 2, line);

    // Restore __js_this. Result is on top of stack — stash it, restore
    // the global, then re-push the result so the caller sees the same
    // shape as before this helper.
    let result_slot = chunks[current].local_count;
    chunks[current].local_count = chunks[current]
        .local_count
        .checked_add(1)
        .expect("emit_invoke_method: local slot overflow");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    c.emit_op_u16(Op::LOCAL_GET, prev_this_slot, line);
    crate::primitives::globals::emit_write(c, "__js_this", line);
    c.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Emit a receiver-once protocol method call.
///
/// This is the language-neutral skeleton for dynamic method syntax such as
/// `receiver.method(args)` / `receiver:method(args)` when a language needs its
/// own property protocol before the final call.
///
/// Stack before: `[receiver, method_key, arg1, ..., argN]`
/// Stack after : `[result]`
///
/// `emit_lookup` receives saved `receiver_slot` and `method_key_slot` and must
/// leave the method/callable value on the stack. `emit_call` receives the stack
/// rebuilt as `[method, receiver, arg1, ..., argN]` and performs the final
/// callable/protocol dispatch.
pub fn emit_protocol_method_call<L, C>(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
    mut emit_lookup: L,
    mut emit_call: C,
) where
    L: FnMut(&mut Vec<Chunk>, usize, u16, u16, u32),
    C: FnMut(&mut Vec<Chunk>, usize, u8, u32),
{
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_string_const("attempt to call a non-function value", line);
        crate::primitives::errors::emit_throw(&mut chunks[current], line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    let method = chunks[current].alloc_scratch(1);
    for i in (0..argc).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
    }

    emit_lookup(chunks, current, base, base + 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, method, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    for i in 2..argc {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
    }
    emit_call(chunks, current, argc, line);
}
