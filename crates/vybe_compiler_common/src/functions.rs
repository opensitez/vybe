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
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

// ── Default parameter handling ──────────────────────────────────────────

/// Emit the start of a default parameter check.
/// If the parameter at `param_slot` is null (missing arg), the caller should compile
/// the default expression, then call `emit_default_param_end`.
/// Returns a jump offset to patch.
/// Stack: unchanged
pub fn emit_default_param_start(chunk: &mut Chunk, param_slot: u16, line: u32) -> usize {
    chunk.emit_op_u16(Op::local_get, param_slot, line);
    chunk.emit_op(Op::ref_is_null, line);
    chunk.emit_jump(Op::br_if_false, line)
}

/// Emit the end of a default parameter check.
/// Caller must have compiled the default expression onto the stack.
/// Stack before: [default_value]  Stack after: [] (stored in param_slot)
pub fn emit_default_param_end(chunk: &mut Chunk, param_slot: u16, skip_jump: usize, line: u32) {
    chunk.emit_op_u16(Op::local_set, param_slot, line);
    chunk.emit_op(Op::drop, line);
    chunk.patch_jump(skip_jump);
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
    chunk.emit_op(Op::null, line);
    chunk.emit_op(Op::r#return, line);
}

/// Emit ref_func to push a closure reference onto the stack.
/// `func_chunk_idx`: the chunk index of the compiled function.
/// `upvalue_count`: 0 for most functions, >0 for closures.
/// Stack: [] → [closure_ref]
pub fn emit_ref_func(chunk: &mut Chunk, func_chunk_idx: usize, upvalue_count: u8, line: u32) {
    chunk.emit_op_u16(Op::ref_func, func_chunk_idx as u16, line);
    chunk.emit(upvalue_count, line);
}

/// Store a function as a global variable.
/// Caller must have closure_ref on stack (from emit_ref_func).
/// Stack before: [closure_ref]  Stack after: []
pub fn emit_store_global_func(chunk: &mut Chunk, name: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::global_set, idx, line);
    chunk.emit_op(Op::drop, line);
}

/// Store a function in a local slot.
/// Caller must have closure_ref on stack (from emit_ref_func).
/// Stack before: [closure_ref]  Stack after: []
pub fn emit_store_local_func(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::local_set, slot, line);
    chunk.emit_op(Op::drop, line);
}

// ── Cross-language function call ────────────────────────────────────────

/// Emit a call to a function by global name.
/// Pushes the function ref, then caller pushes args, then calls emit_call_args.
/// Stack: [] → [function_ref]
pub fn emit_push_global_func(chunk: &mut Chunk, name: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::global_get, idx, line);
}

/// Emit the call opcode after function ref + args are on stack.
/// Stack before: [func_ref, arg1, arg2, ...]  Stack after: [return_value]
pub fn emit_call(chunk: &mut Chunk, arg_count: u8, line: u32) {
    chunk.emit_op_u8(Op::call, arg_count, line);
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
//     2. Emit Op::promise_suspend (WASM JSPI) — VM checks if Promise, suspends fiber if pending
//
// Python `async def`, Dart `async`, JS `async function`, C# `async Task`
// all compile to the same opcodes.

/// Emit an await expression (WASM JSPI: promise_suspend).
/// Caller must have compiled the awaited expression onto the stack.
/// If the value is a Promise, the VM suspends the current fiber until resolved.
/// If the value is not a Promise, it passes through unchanged.
/// Stack before: [value_or_promise]  Stack after: [resolved_value]
pub fn emit_await(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::promise_suspend, line);
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
    chunk.emit_op_u16(Op::ref_func, body_chunk_idx as u16, line);
    chunk.emit(0, line); // 0 upvalues
    chunk.emit_op(Op::cont_new, line);
    // Resume the continuation immediately — it will suspend at each await point
    // The VM's event loop handles re-resumption when promises resolve
    let zero_tag = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::resume, zero_tag, line);
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

/// Emit: create empty args array → push one argument → leave array on stack.
/// Call this for each non-spread argument in a spread call.
/// Stack before: [args_array, value]  Stack after: [args_array]
pub fn emit_spread_push_arg(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_push, line);
}

/// Emit: concat a spread array into the args array.
/// Call this for each spread argument: `...arr`.
/// Stack before: [args_array, spread_array]  Stack after: [merged_array]
pub fn emit_spread_concat_arg(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_concat, line);
}
