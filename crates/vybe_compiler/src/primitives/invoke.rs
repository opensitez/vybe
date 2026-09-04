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
    // The receiver is stashed so it can be pushed as argument 0 of the call
    // below — §10.2.1 `[[Call]](thisArgument, argumentsList)`.
    let c = &mut chunks[current];
    let temp_base = c.alloc_scratch(argc as u16 + 1);
    let receiver_slot = temp_base;
    let arg_base = temp_base + 1;

    // Pop args into temps (LIFO: last arg lands in highest temp slot).
    for i in (0..argc).rev() {
        let slot = arg_base + i as u16;
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    // Stash receiver.
    c.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);

    // ⛔ NOTHING IS PUSHED HERE. This is where the pre-decomposition shape
    // rebuilt `receiver, name, args…` to feed one `emit_call(invoke,
    // argc + 2)`. Decomposing into `getMethodForCall` + `call_ref` gave BOTH
    // arms their own operands, so these pushes fed nothing and left `argc + 2`
    // values on the stack at every dynamic method call.
    //
    // The VM shares one operand stack across blocks, so the leftovers bury
    // whatever the ENCLOSING expression was building. Counted from the
    // disassembly of `[1, ...b.slice(0, 1)]`: the literal's accumulator sits
    // under `b, "slice", 0, 1` by the time the call returns, and the `1` is
    // lost. Invisible in a statement, where leftovers die at the boundary, and
    // invisible to `wasm-tools`, which stops earlier on a continuation type.
    // Only an enclosing expression shows it.

    // ⛔ `invokeMethod(receiver, name, ...args)` IS NOT REPRESENTABLE AS AN
    // IMPORT. Its arity was `argc + 2`, so a different one at every call site,
    // and a WASM import has exactly ONE signature. The writer types host
    // imports by an arity SCAN — `max(argc)` over call sites — so every site
    // with fewer arguments under-supplies, and V8 reports it as "not enough
    // arguments on the stack for call (need 4, got 3)" at the SHORT site,
    // nowhere near the cause. `.map().filter().reduce()` failed exactly there.
    //
    // ⇒ It is also an ECMA defect, and the fix is the spec's own shape.
    // §13.3.6 evaluates `o.m(a, b)` as TWO operations, neither variadic:
    // §7.3.2 `GetV(receiver, key)` then §7.3.14 `Call(F, V, argumentsList)`.
    // `getMethodForCall` is `GetV` and now has a CONSTANT arity of 3; `Call`
    // becomes `call_ref`, which is an INSTRUCTION carrying a real per-arity
    // functype rather than an import, so varying argument counts are expressed
    // by the type system instead of being flattened into one signature.
    //
    // The receiver reaches the body as argument 0 of the `call_ref` below —
    // this changes how the method is REACHED, not how the receiver is passed.
    let lookup = chunks[current].add_import("ecma:value", "getMethodForCall");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    c.emit_string_const(method_name, line);
    // Third argument = "bind the receiver into a host function".
    //
    // ⛔ UNDER `ReceiverAbi::Parameter` THIS MUST BE `false`, AND THE HIT ARM
    // MUST PASS THE RECEIVER ITSELF. Binding is `__bound_args` — a hidden
    // prepend on a funcref-shaped value, which a WASM functype cannot express
    // — so M5 replaces it with a real argument 0.
    //
    // Asking to bind AND passing nothing was the worst of both: the host only
    // binds a host function carrying `__vybe_method_receiver`, so an UNMARKED
    // builtin like `ecma:array.slice` was neither bound nor passed a receiver,
    // and `array_of(args, 0)` read the START INDEX as its array.
    // `[1,2,3,4].slice(1)` returned `[]` and `indexOf` returned -1, while
    // `push`/`concat` — which carry the marker — were fine.
    //
    // The two halves are one decision, which is why they are read from one
    // place: bind XOR pass, never both (that is the `[arr, arr, callback]`
    // duplicate the host already warns about) and never neither.
    let c = &mut chunks[current];
    c.emit_bool_const(false, line);
    c.emit_call(lookup, 3, line);

    // ⛔ A MISS MUST STILL REACH THE HOST. `GetV` here walks USER members —
    // `lookup_user_member_on_chain` — so a BUILTIN method on an intrinsic
    // prototype (`"abc".toUpperCase()`) resolves to null, and calling that is
    // `null is not callable`. The retired `invokeMethod` never had this
    // problem because its host `dispatch` routed builtins directly. Decomposing
    // without keeping a miss path took ~677 js tests with it.
    //
    // So: resolved → `call_ref` (the count rides a real functype); unresolved →
    // the host, in §7.3.14's OTHER shape, with the arguments as ONE list. Both
    // arms leave a value, and BOTH import arities stay constant.
    let method_slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, method_slot, line);
    c.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if_value(line);

    // MISS — hand it to the host so it raises the language's own error, or
    // routes a builtin the user chain cannot see.
    let invoke = chunks[current].add_import("ecma:value", "invokeMethod");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    c.emit_string_const(method_name, line);
    for i in 0..argc {
        c.emit_op_u16(Op::LOCAL_GET, arg_base + i as u16, line);
    }
    crate::primitives::collections::emit_array_new(chunks, current, argc as u16, line);
    let c = &mut chunks[current];
    c.emit_call(invoke, 3, line);

    c.emit_else(line);

    // HIT — `Call(F, V, argumentsList)`: the method is the callee and the
    // arguments follow it, their count carried by `call_ref`'s functype.
    c.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    // §10.2.1 `[[Call]](thisArgument, argumentsList)` — the receiver is
    // argument 0.
    c.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    let recv_argc = 1;
    for i in 0..argc {
        c.emit_op_u16(Op::LOCAL_GET, arg_base + i as u16, line);
    }
    crate::primitives::functions::emit_call(c, argc + recv_argc, line);
    c.emit_end(line);

    // The result is on top of the stack; stash it so the caller sees the same
    // shape this helper has always left.
    let result_slot = chunks[current].local_count;
    chunks[current].local_count = chunks[current]
        .local_count
        .checked_add(1)
        .expect("emit_invoke_method: local slot overflow");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_SET, result_slot, line);
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
