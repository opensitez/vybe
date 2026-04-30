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

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

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
    let temp_base = c.local_count;
    // slots: [receiver, prev_js_this, args...]
    c.local_count = c
        .local_count
        .checked_add(argc as u16 + 2)
        .expect("emit_invoke_method: local slot overflow");
    let receiver_slot = temp_base;
    let prev_this_slot = temp_base + 1;
    let arg_base = temp_base + 2;

    // Pop args into temps (LIFO: last arg lands in highest temp slot).
    for i in (0..argc).rev() {
        let slot = arg_base + i as u16;
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
        c.emit_op(Op::DROP, line);
    }
    // Stash receiver.
    c.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    c.emit_op(Op::DROP, line);

    // Save current __js_this so we can restore after the call. Host
    // functions don't manage this global — every JS method call site
    // (here and in compiler/calls.rs) is responsible for save/restore.
    let js_this_const = c.add_constant(Value::String(Arc::from("__js_this")));
    c.emit_op_u16(Op::GLOBAL_GET, js_this_const, line);
    c.emit_op_u16(Op::LOCAL_SET, prev_this_slot, line);
    c.emit_op(Op::DROP, line);

    // Set __js_this = receiver so JS-compiled method bodies see the
    // right `this` when dispatch eventually drives `ctx.invoke`.
    c.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    c.emit_op_u16(Op::GLOBAL_SET, js_this_const, line);
    c.emit_op(Op::DROP, line);

    // Rebuild call stack: receiver, name, args...
    c.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    let name_const = c.add_constant(Value::String(Arc::from(method_name)));
    c.emit_op_u16(Op::CONST, name_const, line);
    for i in 0..argc {
        let slot = arg_base + i as u16;
        c.emit_op_u16(Op::LOCAL_GET, slot, line);
    }

    let idx = chunks[0].add_import("ecma:value", "invokeMethod");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::CALL_IMPORT, idx, line);
    c.emit(argc + 2, line);

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
    c.emit_op(Op::DROP, line);
    c.emit_op_u16(Op::LOCAL_GET, prev_this_slot, line);
    c.emit_op_u16(Op::GLOBAL_SET, js_this_const, line);
    c.emit_op(Op::DROP, line);
    c.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}
