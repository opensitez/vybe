//! Promise chain emitters — .then(), .catch(), .finally() via WASM JSPI.
//!
//! All languages with async/await, promises, futures, or tasks use the same
//! JSPI mechanism: `call $jspi.await(value)` suspends the fiber until the
//! value settles; the VM's event loop resumes it. This module provides
//! the bytecode recipes for promise chain methods so every language emits
//! identical WASM.
//!
//! Pattern: each chain method (then/catch/finally) is a small bytecode
//! chunk that awaits the input promise, calls the user callback, and
//! returns the result. The chunk uses `emit_await` (JSPI suspend) for
//! the await point — no custom opcodes.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

use super::functions;
use std::sync::Arc;

/// Emit a `.then(onFulfilled, onRejected)` chain step.
/// Stack before: [promise, onFulfilled, onRejected_or_null]
/// Stack after: [result]
///
/// Bytecode pattern:
///   result = await promise
///   if onFulfilled != null: result = onFulfilled(result)
///   return result
/// On exception:
///   if onRejected != null: return onRejected(error)
///   else: rethrow
/// Emit `.then(promise, onFulfilled, onRejected)` via WASM JSPI.
/// Params at slots 0-2. Uses try/catch for rejection handling.
pub fn emit_then(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let input_slot: u16 = 0;
    let fulfilled_slot: u16 = 1;
    let rejected_slot: u16 = 2;
    let result_slot = chunk.alloc_scratch(1);
    let err_slot = chunk.alloc_scratch(1);

    let try_patch = super::errors::emit_try_start(chunk, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    functions::emit_await(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, fulfilled_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, fulfilled_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_end(line);

    super::errors::emit_try_end(chunk, line);
    chunk.emit_op(Op::RETURN, line);

    super::errors::patch_catch(chunk, try_patch);
    chunk.emit_op_u16(Op::LOCAL_SET, err_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rejected_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, rejected_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, err_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, err_slot, line);
    emit_rejected_promise(chunk, line);
    chunk.emit_end(line);
}

/// Build a rejected promise object from the error value on the stack.
/// Stack: [error] → [promise_object]
/// Pure WASM GC — no host calls.
fn emit_rejected_promise(chunk: &mut Chunk, line: u32) {
    let err = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, err, line);
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    let type_key = chunk.add_constant(vybe_bytecode::Value::String(Arc::from("__type")));
    let state_key = chunk.add_constant(vybe_bytecode::Value::String(Arc::from("__state")));
    let value_key = chunk.add_constant(vybe_bytecode::Value::String(Arc::from("__value")));
    chunk.emit_dup(line);
    chunk.emit_string_const("Promise", line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_dup(line);
    chunk.emit_string_const("rejected", line);
    chunk.emit_op_u16(Op::STRUCT_SET, state_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, err, line);
    chunk.emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Emit `.catch(promise, onRejected)` via WASM JSPI.
pub fn emit_catch(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let input_slot: u16 = 0;
    let rejected_slot: u16 = 1;

    let try_patch = super::errors::emit_try_start(chunk, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    functions::emit_await(chunk, line);

    super::errors::emit_try_end(chunk, line);
    chunk.emit_op(Op::RETURN, line);

    super::errors::patch_catch(chunk, try_patch);
    let err_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, err_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rejected_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, rejected_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, err_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, err_slot, line);
    chunk.emit_op(Op::THROW, line);
    chunk.emit_end(line);
}

/// Emit `.finally(promise, onFinally)` via WASM JSPI.
pub fn emit_finally(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let input_slot: u16 = 0;
    let callback_slot: u16 = 1;
    let result_slot = chunk.alloc_scratch(1);
    let err_slot = chunk.alloc_scratch(1);
    let saved_cb = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, saved_cb, line);

    let try_patch = super::errors::emit_try_start(chunk, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    functions::emit_await(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    super::errors::emit_try_end(chunk, line);

    emit_call_if_nonnull(chunk, saved_cb, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::RETURN, line);

    super::errors::patch_catch(chunk, try_patch);
    chunk.emit_op_u16(Op::LOCAL_SET, err_slot, line);
    emit_call_if_nonnull(chunk, saved_cb, line);
    chunk.emit_op_u16(Op::LOCAL_GET, err_slot, line);
    emit_rejected_promise(chunk, line);
}

fn emit_call_if_nonnull(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 0, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
}
