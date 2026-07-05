//! Function compilation helpers — shared bytecode patterns for function scaffolding.
//!
//! Every language compiles functions the same way at the bytecode level:
//! - Create a Chunk (name, arity)
//! - Map params to local slots
//! - Handle default values
//! - Compile body (language-specific)
//! - Emit null + return as safety net
//! - Store ref_func as local/global
//!
//! The scaffolding (everything except body compilation) is identical.
//! Python `def`, Dart `void f()`, JS `function`, C# `void F()`, VB `Sub`
//! all produce the same Chunk structure.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

// ── Default parameter handling ──────────────────────────────────────────

/// Emit the start of a default parameter check.
/// If the parameter at `param_slot` is null (missing arg), the caller should compile
/// the default expression, then call `emit_default_param_end`.
/// Returns a structured block patch to close.
/// Stack: unchanged
pub fn emit_default_param_start(chunk: &mut Chunk, param_slot: u16, line: u32) -> usize {
    chunk.emit_op_u16(Op::LOCAL_GET, param_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let block = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    block
}

/// Emit the end of a default parameter check.
/// Caller must have compiled the default expression onto the stack.
/// Stack before: [default_value]  Stack after: [] (stored in param_slot)
pub fn emit_default_param_end(chunk: &mut Chunk, param_slot: u16, block: usize, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, param_slot, line);
    chunk.emit_end(line);
    chunk.patch_block(block);
}

// ── Function chunk scaffolding ──────────────────────────────────────────

/// Create a new function chunk with the given name and arity.
/// Returns the chunk — caller adds it to their chunks vec and manages the scope.
pub fn create_function_chunk(name: &str, arity: u8) -> Chunk {
    let mut chunk = Chunk::new(name);
    chunk.arity = arity;
    chunk
}

/// Emit the function epilogue: null return (safety net for functions that
/// fall through without explicit return).
/// Stack: [] → diverges (return)
pub fn emit_function_epilogue(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::NULL, line);
    chunk.emit_op(Op::RETURN, line);
}

/// Start shared async-body scaffolding for function-like chunks.
/// Returns the catch jump to patch once the body has been emitted.
/// The caller remains responsible for Promise.resolve / Promise.reject
/// wrapping because import indices live at the compiler layer.
pub fn emit_async_body_start(chunk: &mut Chunk, line: u32) -> usize {
    crate::errors::emit_try_start(chunk, line)
}

/// Finish the normal fallthrough path of a shared async body.
/// Leaves `undefined` on the stack so the compiler can wrap it with
/// `Promise.resolve(undefined)` before returning.
pub fn emit_async_body_fallthrough(chunk: &mut Chunk, catch_jump: usize, line: u32) {
    crate::errors::emit_try_end(chunk, line);
    chunk.emit_op(Op::NULL, line);
    let _ = catch_jump;
}

/// Patch the catch edge for a shared async body so the compiler can
/// emit its rejection path.
pub fn patch_async_body_catch(chunk: &mut Chunk, catch_jump: usize) {
    crate::errors::patch_catch(chunk, catch_jump);
}

/// Emit ref_func to push a closure reference onto the stack.
/// `func_chunk_idx`: the chunk index of the compiled function.
/// `upvalue_count`: 0 for most functions, >0 for closures.
/// Stack: [] → [closure_ref]
pub fn emit_ref_func(chunk: &mut Chunk, func_chunk_idx: usize, upvalue_count: u8, line: u32) {
    chunk.emit_op_u16(Op::REF_FUNC, func_chunk_idx as u16, line);
    chunk.emit(upvalue_count, line);
}

/// Store a function as a global variable.
/// Caller must have closure_ref on stack (from emit_ref_func).
/// Stack before: [closure_ref]  Stack after: []
pub fn emit_store_global_func(chunk: &mut Chunk, name: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::GLOBAL_SET, idx, line);
}

/// Store a function in a local slot.
/// Caller must have closure_ref on stack (from emit_ref_func).
/// Stack before: [closure_ref]  Stack after: []
pub fn emit_store_local_func(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

// ── Cross-language function call ────────────────────────────────────────

/// Emit a call to a function by global name.
/// Pushes the function ref, then caller pushes args, then calls emit_call_args.
/// Stack: [] → [function_ref]
pub fn emit_push_global_func(chunk: &mut Chunk, name: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::GLOBAL_GET, idx, line);
}

/// Emit the call opcode after function ref + args are on stack.
/// Stack before: [func_ref, arg1, arg2, ...]  Stack after: [return_value]
pub fn emit_call(chunk: &mut Chunk, arg_count: u8, line: u32) {
    chunk.emit_op_u8(Op::CALL_REF, arg_count, line);
}

// ── Async/await (WASM Stack Switching + JSPI) ───────────────────────────
//
// All languages use the same async pattern:
//
//   async function:
//     1. Create continuation from body function (cont_new)
//     2. Return the continuation as a Promise-like value
//     3. The runtime schedules it on the event loop
//
//   await expression:
//     1. Compile the expression (produces a value or Promise)
//     2. Emit the spec stack-switching `suspend` instruction tagged with
//        AWAIT_SUSPEND_TAG (JSPI is stack switching applied to JS Promises) —
//        the VM checks if it's a Promise and suspends the fiber if pending.
//        No custom opcode is involved.
//
// Python `async def`, Dart `async`, JS `async function`, C# `async Task`
// all compile to the same opcodes.

/// Module/name of the JSPI suspending import that `await` lowers to. JSPI
/// (JS Promise Integration — stack switching at the JS-promise boundary) marks
/// an import as `WebAssembly.Suspending`; calling it suspends the computation
/// until the returned Promise settles. The VM recognises this import as the
/// suspender (the embedder-side marking) and runs the await/suspend logic.
pub const JSPI_SUSPEND_MODULE: &str = "jspi";
pub const JSPI_SUSPEND_NAME: &str = "await";

/// Emit an `await` expression via JSPI — a plain `call` to a suspending import.
/// Caller must have compiled the awaited expression onto the stack.
/// Stack before: `[value_or_promise]`  Stack after: `[resolved_value]`
pub fn emit_await(chunk: &mut Chunk, line: u32) {
    // Per the JSPI proposal, the suspend point is a normal `call` to a
    // `WebAssembly.Suspending`-marked import — NOT a custom opcode and NOT a
    // magic `suspend` tag. `await x` → `call $jspi.await(x)`, which lowers to
    // the core `call` (0x10) — valid `.wasm`. The VM treats this import as the
    // suspender: fulfilled → unwrap, rejected → throw, pending → suspend the
    // fiber on the event loop (the engine-internal stack switch JSPI mandates)
    // until the Promise settles, then resume with its value. A non-Promise
    // value passes straight through (proposal §"nosuspend").
    let idx = chunk.add_import(JSPI_SUSPEND_MODULE, JSPI_SUSPEND_NAME);
    chunk.emit_call(idx, 1, line); // argc = 1 (the awaited value)
}

/// Two-chunk `await` for runtime-helper builders: the awaited-value `call` is
/// emitted into `code`, but the `jspi.await` import is registered on `imports`
/// (chunks[0]) — matching how those builders register every other import.
/// Adding it to `code`'s own import list instead would shift `code`'s import
/// indices and mis-resolve its other `CALL_IMPORT`s.
pub fn emit_await_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let idx = code.add_import(JSPI_SUSPEND_MODULE, JSPI_SUSPEND_NAME);
    code.emit_call(idx, 1, line); // argc = 1 (the awaited value)
}

/// Emit async function wrapper: wraps the body chunk as a continuation.
/// Call this INSTEAD of the normal function body compilation for async functions.
///
/// The pattern:
///   1. The outer function creates a continuation from the body chunk
///   2. Returns a Promise that resolves when the continuation completes
///
/// `body_chunk_idx`: the chunk index containing the compiled async body
/// Stack: [] → [promise]
pub fn emit_async_wrapper(chunk: &mut Chunk, body_chunk_idx: usize, line: u32) {
    // Create continuation from the body function
    chunk.emit_op_u16(Op::REF_FUNC, body_chunk_idx as u16, line);
    chunk.emit(0, line); // 0 upvalues
    chunk.emit_op(Op::CONT_NEW, line);
    // Resume the continuation immediately — it will suspend at each await point
    // The VM's event loop handles re-resumption when promises resolve
    crate::generators::emit_resume(chunk, line);
}

/// Create an async function body chunk.
/// Same as create_function_chunk but named with $async suffix for debugging.
pub fn create_async_body_chunk(name: &str, arity: u8) -> Chunk {
    create_function_chunk(&format!("{}$async", name), arity)
}

// ── Spread arguments ───────────────────────────────────────────────────
//
// When a call has spread arguments: f(a, ...arr, b)
// The compiler builds an args array at runtime:
//   1. array_new 0 (empty array)
//   2. For each normal arg: compile + array_push
//   3. For each spread arg: compile + array_concat (flattens into the array)
//   4. Use the array length as argc for the call
//
// This is language-agnostic — JS, Python (*args), Ruby (*splat) all use this.

/// Emit: push one argument onto a spread-args array.
/// Stack before: [args_array, value]  Stack after: [args_array]
///
/// Routes through `ecma:array.push` (returns new length per
/// ECMA-262); caller stashes arr in a local before the loop and
/// reloads afterwards — see `compile_function_decl` rest-args for
/// the canonical template. This helper assumes caller has the stack
/// preserved via a local.
pub fn emit_spread_push_arg(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::collections::emit_push(chunks, current, line);
}

/// Emit: concat a spread array into the args array via
/// `ecma:array.concat` — returns a new array; caller replaces
/// the accumulator local with the result.
pub fn emit_spread_concat_arg(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::collections::emit_concat(chunks, current, line);
}
