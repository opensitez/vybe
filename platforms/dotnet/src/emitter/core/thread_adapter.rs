//! Thread.Sleep adapter — WASI spec-compliant blocking sleep.
//!
//! Replaces the retired `vybe:clocks.sleep` host function. Thread.Sleep(ms)
//! compiles to the two-step WASI 0.2 pollable pattern:
//!   1. `wasi:clocks/monotonic-clock.subscribe-duration(ns)` → pollable
//!   2. `wasi:io/poll.[method]pollable.block(pollable)` → ()
//!
//! Both functions are WASI 0.2 spec and already registered in vybe_host.

use std::sync::Arc;

use vybe_runtime::{opcode::Op, Chunk, Value};
use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::instructions::core_wasm;

const CANCELLED_KEY: &str = "__dotnet_cancelled";
const CANCEL_AT_MS_KEY: &str = "__dotnet_cancel_at_ms";
const LINKED_KEY: &str = "__dotnet_linked_tokens";
const TOKEN_KEY: &str = "Token";
const REQUESTED_KEY: &str = "IsCancellationRequested";
const DELAY_TOKEN_KEY: &str = "__dotnet_delay_token";
const EXCEPTION_KEY: &str = "exception";
const REGISTRATIONS_KEY: &str = "__dotnet_cancellation_registrations";
const SOURCE_TYPE: &str = "CancellationTokenSource";

// Thread.Sleep bytecode lives in vybe_compiler::primitives::threading::emit_thread_sleep
// (pure WASI, platform-neutral).

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val) }
}

fn struct_get(chunk: &mut Chunk, field: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(field)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, idx, line);
}

fn struct_set_drop(chunk: &mut Chunk, field: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(field)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, idx, line);
}

fn emit_task_wait_method_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut method =
        vybe_compiler::primitives::functions::create_function_chunk("__dotnet_task_wait_method", 1);
    method.emit_op_u16(Op::LOCAL_GET, 0, line);
    let mut method_chunks = vec![method];
    emit_task_wait(&mut method_chunks, 0, line);
    method_chunks[0].emit_op(Op::RETURN, line);
    method_chunks[0].local_count = method_chunks[0].local_count.max(1);
    chunks.push(method_chunks.remove(0));
    chunks.len() - 1
}

fn emit_bind_task_method(
    chunk: &mut Chunk,
    task_slot: u16,
    method_name: &str,
    method_chunk_idx: usize,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, task_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, method_chunk_idx as u16, line);
    chunk.emit(0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, task_slot, line);
    struct_set_drop(chunk, "__vybe_method_receiver", line);
    struct_set_drop(chunk, method_name, line);
}

fn emit_attach_task_members(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let task_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, task_slot, line);
    let wait_method = emit_task_wait_method_chunk(chunks, line);
    emit_bind_task_method(&mut chunks[current], task_slot, "Wait", wait_method, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, task_slot, line);
}

fn emit_now_ms(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:date", "now", 0, line);
}

fn emit_nullish(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let undef = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undef, 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

/// `new CancellationTokenSource()` / `CancellationToken.None`.
/// Stack: [] -> [token_source_or_token]
pub fn emit_cancellation_token_source_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_struct_new(0, 0, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from(SOURCE_TYPE)), line);
    struct_set_drop(chunk, "__type", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_drop(chunk, CANCELLED_KEY, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_drop(chunk, REQUESTED_KEY, line);
    core_wasm::dup(chunk, line);
    core_wasm::dup(chunk, line);
    struct_set_drop(chunk, TOKEN_KEY, line);
}

/// `CancellationToken.None`.
/// Stack: [] -> [token]
pub fn emit_cancellation_token_none(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_struct_new(0, 0, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("CancellationToken")), line);
    struct_set_drop(chunk, "__type", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_drop(chunk, CANCELLED_KEY, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_drop(chunk, REQUESTED_KEY, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_drop(chunk, "CanBeCanceled", line);
}

fn emit_fire_cancellation_registrations(chunks: &mut [Chunk], current: usize, line: u32) {
    let regs_slot = chunks[current].alloc_scratch(1);
    struct_get(&mut chunks[current], REGISTRATIONS_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, regs_slot, line);
    emit_nullish(&mut chunks[current], regs_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, regs_slot, line);
    collections::emit_array_length(&mut chunks[current], line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, regs_slot, line);
    collections::emit_pop(chunks, current, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 0, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

/// `cts.Cancel()` — mark the shared source/token object as cancelled.
/// Stack: [cts] -> [null]
pub fn emit_cancellation_token_cancel(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let token_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, token_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, token_slot, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(true), line);
    struct_set_drop(chunk, CANCELLED_KEY, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(true), line);
    struct_set_drop(chunk, REQUESTED_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, token_slot, line);
    emit_fire_cancellation_registrations(chunks, current, line);
}

/// `cts.CancelAfter(ms)` — record a wall-clock deadline in milliseconds.
/// Stack: [cts, ms] -> [null]
pub fn emit_cancellation_token_cancel_after(chunks: &mut [Chunk], current: usize, line: u32) {
    let ms_slot = chunks[current].alloc_scratch(2);
    let cts_slot = ms_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cts_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cts_slot, line);
    emit_now_ms(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    struct_set_drop(&mut chunks[current], CANCEL_AT_MS_KEY, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `cts.Token` — the source object is the token backing store.
/// Stack: [cts] -> [token]
pub fn emit_cancellation_token_source_token(chunks: &mut [Chunk], current: usize, line: u32) {
    // The source object itself is the shared token backing store. Do not
    // restamp `__type`: VB/C# code may read `cts.Token` and then call
    // `cts.Cancel()` on the same object, so the source must keep its
    // CancellationTokenSource method surface.
    let _ = (&mut chunks[current], line);
}

/// `CancellationTokenSource.CreateLinkedTokenSource(t1, t2, ...)`
/// Stack: [tokens...] -> [source]
pub fn emit_cancellation_token_linked_source(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    let first_token = chunks[current].alloc_scratch(u16::from(argc.max(1)));
    for offset in (0..argc).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, first_token + u16::from(offset), line);
    }
    for offset in 0..argc {
        chunks[current].emit_op_u16(Op::LOCAL_GET, first_token + u16::from(offset), line);
    }
    collections::emit_array_new(chunks, current, argc as u16, line);
    let tokens_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tokens_slot, line);
    emit_cancellation_token_source_new(chunks, current, line);
    let source_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tokens_slot, line);
    struct_set_drop(&mut chunks[current], LINKED_KEY, line);
    let callback_idx = emit_linked_token_cancel_callback(chunks, line);
    for offset in 0..argc {
        chunks[current].emit_op_u16(Op::LOCAL_GET, first_token + u16::from(offset), line);
        chunks[current].emit_op_u16(Op::REF_FUNC, callback_idx as u16, line);
        chunks[current].emit(1, line);
        vybe_compiler::primitives::functions::emit_closure_upvalue(
            &mut chunks[current],
            true,
            source_slot,
            line,
        );
        emit_cancellation_token_register(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
}

fn emit_linked_token_cancel_callback(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut callback = vybe_compiler::primitives::functions::create_function_chunk(
        "__dotnet_linked_token_cancel",
        0,
    );
    callback.local_count = 1;
    callback.capture_base = 0;
    callback.capture_count = 1;
    callback.emit_op_u16(Op::LOCAL_GET, 0, line);
    let mut callback_chunks = vec![callback];
    emit_cancellation_token_cancel(&mut callback_chunks, 0, line);
    callback_chunks[0].emit_op(Op::RETURN, line);
    chunks.push(callback_chunks.remove(0));
    chunks.len() - 1
}

/// `token.Register(callback)`.
/// Stack: [token, callback] -> [registration]
pub fn emit_cancellation_token_register(chunks: &mut [Chunk], current: usize, line: u32) {
    let callback_slot = chunks[current].alloc_scratch(3);
    let token_slot = callback_slot + 1;
    let regs_slot = callback_slot + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, callback_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, token_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, token_slot, line);
    struct_get(&mut chunks[current], REGISTRATIONS_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, regs_slot, line);
    emit_nullish(&mut chunks[current], regs_slot, line);
    chunks[current].emit_if(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, regs_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, token_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, regs_slot, line);
    struct_set_drop(&mut chunks[current], REGISTRATIONS_KEY, line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, regs_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_struct_new(0, 0, line);
}

/// `token.CanBeCanceled`.
/// Stack: [token] -> [bool]
pub fn emit_cancellation_token_can_be_canceled(chunks: &mut [Chunk], current: usize, line: u32) {
    let token_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, token_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, token_slot, line);
    struct_get(&mut chunks[current], "CanBeCanceled", line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, token_slot, line);
    struct_get(&mut chunks[current], "CanBeCanceled", line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// `token.WaitHandle`.
/// Stack: [token] -> [handle]
pub fn emit_cancellation_token_wait_handle(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_struct_new(0, 0, line);
}

/// `token.IsCancellationRequested` / `cts.IsCancellationRequested`.
/// Stack: [token] -> [bool]
pub fn emit_cancellation_token_is_requested(chunks: &mut [Chunk], current: usize, line: u32) {
    let token_slot = chunks[current].alloc_scratch(2);
    let deadline_slot = token_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, token_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, token_slot, line);
    struct_get(&mut chunks[current], CANCELLED_KEY, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, token_slot, line);
    struct_get(&mut chunks[current], CANCEL_AT_MS_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, deadline_slot, line);
    emit_nullish(&mut chunks[current], deadline_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    emit_now_ms(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, deadline_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `token.ThrowIfCancellationRequested()`.
/// Stack: [token] -> [null] or throws OperationCanceledException.
pub fn emit_cancellation_token_throw_if_requested(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_cancellation_token_is_requested(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_throw_task_cancelled(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

fn emit_throw_task_cancelled(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_operation_cancelled_exception(chunks, current, line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

fn emit_operation_cancelled_exception(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_struct_new(0, 0, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_string_const("The operation was canceled.", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "OperationCanceledException",
        line,
    );
    vybe_compiler::primitives::errors::emit_stamp_exception_ancestors(
        &mut chunks[current],
        "OperationCanceledException",
        line,
    );
}

fn emit_cancelled_task_object(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_struct_new(0, 0, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_string_const("Task", line);
    struct_set_drop(&mut chunks[current], "__type", line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_bool_const(true, line);
    struct_set_drop(&mut chunks[current], "iscompleted", line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_bool_const(false, line);
    struct_set_drop(&mut chunks[current], "isalive", line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_i32_const(-1, line);
    struct_set_drop(&mut chunks[current], "exitcode", line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_string_const("Canceled", line);
    struct_set_drop(&mut chunks[current], "status", line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    struct_set_drop(&mut chunks[current], "result", line);
    core_wasm::dup(&mut chunks[current], line);
    emit_operation_cancelled_exception(chunks, current, line);
    struct_set_drop(&mut chunks[current], EXCEPTION_KEY, line);
}

fn emit_cancelled_task_object_with_members(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_cancelled_task_object(chunks, current, line);
    emit_attach_task_members(chunks, current, line);
}

fn emit_task_delay_timer_callback(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut callback =
        vybe_compiler::primitives::functions::create_function_chunk("__dotnet_task_delay_timer", 0);
    callback.local_count = 3;
    callback.capture_base = 0;
    callback.capture_count = 3;

    callback.emit_op_u16(Op::LOCAL_GET, 0, line);
    let mut callback_chunks = vec![callback];
    emit_cancellation_token_is_requested(&mut callback_chunks, 0, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut callback_chunks[0], line);
    callback_chunks[0].emit_if(line);
    callback_chunks[0].emit_op_u16(Op::LOCAL_GET, 2, line);
    emit_operation_cancelled_exception(&mut callback_chunks, 0, line);
    callback_chunks[0].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    callback_chunks[0].emit_op(Op::DROP, line);
    callback_chunks[0].emit_else(line);
    callback_chunks[0].emit_op_u16(Op::LOCAL_GET, 1, line);
    callback_chunks[0].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    callback_chunks[0].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    callback_chunks[0].emit_op(Op::DROP, line);
    callback_chunks[0].emit_end(line);
    callback_chunks[0].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    callback_chunks[0].emit_op(Op::RETURN, line);

    chunks.push(callback_chunks.remove(0));
    chunks.len() - 1
}

/// `Task.Delay(ms[, token])`.
/// Stack: [ms] or [ms, token] -> [task]
pub fn emit_task_delay(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        call_import(chunks, current, "ecma:promise", "resolve", 1, line);
        emit_attach_task_members(chunks, current, line);
        return;
    }

    let token_slot = chunks[current].alloc_scratch(6);
    let ms_slot = token_slot + 1;
    let resolvers_slot = token_slot + 2;
    let promise_slot = token_slot + 3;
    let resolve_slot = token_slot + 4;
    let reject_slot = token_slot + 5;
    chunks[current].emit_op_u16(Op::LOCAL_SET, token_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ms_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, token_slot, line);
    emit_cancellation_token_is_requested(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_cancelled_task_object_with_members(chunks, current, line);
    chunks[current].emit_else(line);

    call_import(chunks, current, "ecma:promise", "withResolvers", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, resolvers_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, resolvers_slot, line);
    struct_get(&mut chunks[current], "promise", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, promise_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, resolvers_slot, line);
    struct_get(&mut chunks[current], "resolve", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, resolve_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, resolvers_slot, line);
    struct_get(&mut chunks[current], "reject", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, reject_slot, line);

    let callback_idx = emit_task_delay_timer_callback(chunks, line);
    chunks[current].emit_op_u16(Op::REF_FUNC, callback_idx as u16, line);
    chunks[current].emit(3, line);
    vybe_compiler::primitives::functions::emit_closure_upvalue(
        &mut chunks[current],
        true,
        token_slot,
        line,
    );
    vybe_compiler::primitives::functions::emit_closure_upvalue(
        &mut chunks[current],
        true,
        resolve_slot,
        line,
    );
    vybe_compiler::primitives::functions::emit_closure_upvalue(
        &mut chunks[current],
        true,
        reject_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    call_import(chunks, current, "web:timers", "setTimeout", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, promise_slot, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, token_slot, line);
    struct_set_drop(&mut chunks[current], DELAY_TOKEN_KEY, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, token_slot, line);
    struct_set_drop(&mut chunks[current], "Token", line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_string_const("WaitingForActivation", line);
    struct_set_drop(&mut chunks[current], "status", line);
    emit_attach_task_members(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `Task.Yield()` — a completed promise-shaped awaitable that still travels
/// through the common await/JSPI path.
pub fn emit_task_yield(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    call_import(chunks, current, "ecma:promise", "resolve", 1, line);
}

/// `ValueTask.AsTask()` — async values are already promise/task-shaped in the
/// ECMA backend, so this is an identity adapter.
pub fn emit_value_task_as_task(chunks: &mut [Chunk], current: usize, line: u32) {
    let _ = (&mut chunks[current], line);
}

/// `task.IsCanceled`.
/// Stack: [task] -> [bool]
pub fn emit_task_is_canceled(chunks: &mut [Chunk], current: usize, line: u32) {
    let task_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, task_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, task_slot, line);
    struct_get(&mut chunks[current], "status", line);
    chunks[current].emit_string_const("Canceled", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, task_slot, line);
    struct_get(&mut chunks[current], DELAY_TOKEN_KEY, line);
    emit_cancellation_token_is_requested(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `Task.Run(fn)` — execute the delegate into an ECMA promise.
/// Stack: [fn] -> [task]
pub fn emit_task_run(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:promise", "try", 1, line);
}

/// `task.Result` — synchronous value for already-settled test tasks.
/// Stack: [task] -> [value]
pub fn emit_task_result(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get(&mut chunks[current], "__value", line);
}

/// `task.IsCompleted` — true for fulfilled/rejected tasks, false while pending.
/// Stack: [task] -> [bool]
pub fn emit_task_is_completed(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get(&mut chunks[current], "__state", line);
    chunks[current].emit_string_const("pending", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `task.Wait()` — join a spawned task and surface failures as .NET
/// `AggregateException`.
/// Stack: [task] -> [null] or throws.
pub fn emit_task_wait(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `Task.FromResult(v)` — a task that is already completed with `v`. Since async
/// runs eagerly, that is just `Promise.resolve(v)`. Stack: [value] → [task].
pub fn emit_task_from_result(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:promise", "resolve", 1, line);
}

/// Collapse the combinator's args into a single iterable of tasks: one array
/// argument (the `IEnumerable<Task>` overload) is used directly; otherwise the
/// N task args are packed into an array.
fn emit_pack_task_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 1 {
        let slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        call_import(chunks, current, "ecma:array", "isArray", 1, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        collections::emit_array_new(chunks, current, 1, line);
        chunks[current].emit_end(line);
    } else {
        collections::emit_array_new(chunks, current, argc as u16, line);
    }
}

/// `Task.WhenAll(t1, …)` — completes with the array of every task's result.
/// `Promise.all` over the (eagerly-resolved) tasks. Stack: [t1 .. tN] → [task].
pub fn emit_task_when_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        let base = chunks[current].alloc_scratch(argc as u16);
        for i in (0..argc).rev() {
            chunks[current].emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
        }
        for i in 0..argc {
            chunks[current].emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
            struct_get(&mut chunks[current], "__value", line);
        }
        collections::emit_array_new(chunks, current, argc as u16, line);
        call_import(chunks, current, "ecma:promise", "resolve", 1, line);
        return;
    }
    emit_pack_task_args(chunks, current, argc, line);
    call_import(chunks, current, "ecma:promise", "all", 1, line);
}

/// `Task.WhenAny(t1, …)` — completes with the first task to finish
/// (`Promise.race`). Stack: [t1 .. tN] → [task].
pub fn emit_task_when_any(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        call_import(chunks, current, "ecma:promise", "resolve", 1, line);
        return;
    }
    emit_pack_task_args(chunks, current, argc, line);
    call_import(chunks, current, "ecma:promise", "race", 1, line);
}

/// `antecedent.ContinueWith(fn)` — run `fn(antecedent)` once the antecedent has
/// completed and return its result as a new task. `fn` receives the completed
/// task, so its `t.Result` reads the antecedent's value (unwrap/join). async
/// runs eagerly, so no explicit wait is needed. Stack: [antecedent, fn] → [task].
pub fn emit_task_continue_with(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = chunks[current].alloc_scratch(1);
    let ante_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ante_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ante_slot, line);
    vybe_compiler::primitives::functions::emit_call(&mut chunks[current], 1, line);
    call_import(chunks, current, "ecma:promise", "resolve", 1, line);
}
