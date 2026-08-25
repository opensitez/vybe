//! JVM `java.util.concurrent` executor adapters.
//!
//! `ExecutorService`, `Executors`, and `Future` on the SAME thread machinery
//! `Thread.start()` rides (`list_adapter::emit_java_thread_start_with`):
//! `submit`/`execute` spawn a real wasi thread per task and keep the task
//! handle; `Future.get()` is a join that caches the task's return value.
//! Pool sizing is accepted and irrelevant — every task gets a thread, which
//! is an execution the `ExecutorService` contract permits and the only
//! honest one when the host thread pool IS the OS scheduler.

use std::sync::Arc;
use vybe_compiler::primitives::{
    callable, collections,
    functions::create_function_chunk,
    globals,
    instructions::{core_wasm, host},
    ops, threading,
};
use vybe_runtime::Chunk;
use vybe_runtime::Value;
use vybe_runtime::opcode::Op;

const SHUTDOWN: &str = "__exec_shutdown";
const TASKS: &str = "__exec_tasks";
const FUT_HANDLE: &str = "__future_handle";
const FUT_RECORD: &str = "__future_record";
const FUT_VALUE: &str = "__future_value";
const FUT_DONE: &str = "__future_done";
const REC_RESULT: &str = "__result";
const REC_INVOKE: &str = "__invoke";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn prop_get(chunks: &mut [Chunk], current: usize, obj: u16, key: &str, line: u32) {
    get(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(key, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

fn prop_set_from_stack(chunks: &mut [Chunk], current: usize, obj: u16, key: &str, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(key, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `Executors.newFixedThreadPool(n)` / `newCachedThreadPool()` /
/// `newSingleThreadExecutor()` — the executor object.
pub fn emit_executor_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    host::emit(&mut chunks[current], "ecma:object", "new", 0, line);
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    chunks[current].emit_bool_const(false, line);
    prop_set_from_stack(chunks, current, obj, SHUTDOWN, line);
    collections::emit_array_new(chunks, current, 0, line);
    prop_set_from_stack(chunks, current, obj, TASKS, line);
    get(&mut chunks[current], obj, line);
}

/// The worker chunk every submitted task runs on: receives the task record's
/// table index, publishes itself as the current thread, invokes the record's
/// `__invoke` callable and writes the value into the record's `__result` —
/// the join itself only reports status (0/1), so the RECORD is how the value
/// crosses back. `emit_submit` resolves WHAT to invoke (closure, or the
/// bound `call`/`run` method of an object payload) before the crossing, so
/// the worker has exactly one job.
fn task_worker_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut worker = create_function_chunk("__java_executor_worker", 1);
    let invoke_key = worker.add_constant(Value::String(Arc::from(REC_INVOKE)));
    let result_key = worker.add_constant(Value::String(Arc::from(REC_RESULT)));
    let record = 0u16;
    let value = 1u16;
    worker.emit_op_u16(Op::LOCAL_GET, record, line);
    worker.emit_op_u16(Op::TABLE_GET, 0, line);
    worker.emit_op_u16(Op::LOCAL_SET, record, line);
    worker.emit_op_u16(Op::LOCAL_GET, record, line);
    globals::emit_write(&mut worker, "__j_current_thread", line);
    worker.emit_op_u16(Op::LOCAL_GET, record, line);
    worker.emit_struct_field_op(Op::STRUCT_GET, 0, invoke_key, line);
    callable::emit_direct_invoke_chunk(&mut worker, 0, line);
    worker.emit_op_u16(Op::LOCAL_SET, value, line);
    worker.emit_op_u16(Op::LOCAL_GET, record, line);
    worker.emit_op_u16(Op::LOCAL_GET, value, line);
    worker.emit_struct_field_op(Op::STRUCT_SET, 0, result_key, line);
    worker.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    worker.emit_op(Op::RETURN, line);
    worker.local_count = 2;
    chunks.push(worker);
    chunks.len() - 1
}

/// `pool.submit(task)` → a Future. `pool.execute(task)` → null. Both spawn.
pub fn emit_submit(chunks: &mut Vec<Chunk>, current: usize, returns_future: bool, line: u32) {
    let task = chunks[current].alloc_scratch(1);
    let pool = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], task, line);
    set(&mut chunks[current], pool, line);

    // Resolve WHAT the worker will invoke, here in the submitting frame:
    // an object payload's own `call` (Callable) or `run` (Runnable) method,
    // receiver-bound by the shared lookup; anything without either is a
    // closure and is invoked as itself. A method lookup MISS is Null for
    // object receivers and Undefined for primitives, hence the truthiness
    // tests.
    let invoke = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], task, line);
    chunks[current].emit_string_const("call", line);
    host::emit(&mut chunks[current], "ecma:value", "getMethodForCall", 2, line);
    set(&mut chunks[current], invoke, line);
    get(&mut chunks[current], invoke, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], task, line);
    chunks[current].emit_string_const("run", line);
    host::emit(&mut chunks[current], "ecma:value", "getMethodForCall", 2, line);
    set(&mut chunks[current], invoke, line);
    get(&mut chunks[current], invoke, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], task, line);
    set(&mut chunks[current], invoke, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    // The record that crosses the thread boundary via table 0.
    let record = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    set(&mut chunks[current], record, line);
    get(&mut chunks[current], record, line);
    get(&mut chunks[current], invoke, line);
    {
        let k = chunks[current].add_constant(Value::String(Arc::from(REC_INVOKE)));
        chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
    }

    let worker_idx = task_worker_chunk(chunks, line);
    get(&mut chunks[current], record, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::TABLE_GROW, 0, line);
    chunks[current].emit_op_u16(Op::REF_FUNC, worker_idx as u16, line);
    chunks[current].emit(0, line);
    threading::emit_thread_spawn(chunks, current, line);
    let handle = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], handle, line);

    // Remember the handle so shutdown/awaitTermination can drain the pool.
    prop_get(chunks, current, pool, TASKS, line);
    get(&mut chunks[current], handle, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    if returns_future {
        host::emit(&mut chunks[current], "ecma:object", "new", 0, line);
        let fut = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], fut, line);
        get(&mut chunks[current], handle, line);
        prop_set_from_stack(chunks, current, fut, FUT_HANDLE, line);
        get(&mut chunks[current], record, line);
        prop_set_from_stack(chunks, current, fut, FUT_RECORD, line);
        chunks[current].emit_bool_const(false, line);
        prop_set_from_stack(chunks, current, fut, FUT_DONE, line);
        get(&mut chunks[current], fut, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
}

/// `future.get()` — join once, cache, answer from the cache after.
pub fn emit_future_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let fut = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fut, line);
    prop_get(chunks, current, fut, FUT_DONE, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    prop_get(chunks, current, fut, FUT_VALUE, line);
    chunks[current].emit_else(line);
    prop_get(chunks, current, fut, FUT_HANDLE, line);
    threading::emit_thread_join(&mut chunks[current], line);
    // Join's status: 0 ok / 1 faulted. A task that threw surfaces on
    // `get()` as the JDK's ExecutionException.
    chunks[current].emit_if(line);
    crate::emitter::exceptions::emit_jvm_exception_throw(
        chunks,
        current,
        "ExecutionException",
        line,
    );
    chunks[current].emit_end(line);
    // The value sits in the record's `__result`.
    prop_get(chunks, current, fut, FUT_RECORD, line);
    let record = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], record, line);
    get(&mut chunks[current], record, line);
    {
        let k = chunks[current].add_constant(Value::String(Arc::from(REC_RESULT)));
        chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
    }
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    prop_set_from_stack(chunks, current, fut, FUT_VALUE, line);
    chunks[current].emit_bool_const(true, line);
    prop_set_from_stack(chunks, current, fut, FUT_DONE, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

/// `future.isDone()` / `future.isCancelled()`.
pub fn emit_future_is_done(chunks: &mut [Chunk], current: usize, line: u32) {
    let fut = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fut, line);
    prop_get(chunks, current, fut, FUT_DONE, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_future_is_cancelled(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(false, line);
}

pub fn emit_future_cancel(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc.max(1) {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_bool_const(false, line);
}

/// Drain every remembered task handle (join), leaving the pool terminated.
fn drain_tasks(chunks: &mut [Chunk], current: usize, pool: u16, line: u32) {
    prop_get(chunks, current, pool, TASKS, line);
    let tasks = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], tasks, line);
    let len = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], tasks, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    let i = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], i, line);
    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], tasks, line);
    get(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    threading::emit_thread_join(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
}

/// `pool.shutdown()` — no new tasks; running ones finish (drained here so a
/// later `isTerminated`/`get` sees completed work). → null.
pub fn emit_shutdown(chunks: &mut [Chunk], current: usize, line: u32) {
    let pool = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], pool, line);
    chunks[current].emit_bool_const(true, line);
    prop_set_from_stack(chunks, current, pool, SHUTDOWN, line);
    drain_tasks(chunks, current, pool, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `pool.shutdownNow()` → the (empty) list of never-started tasks.
pub fn emit_shutdown_now(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_shutdown(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    collections::emit_array_new(chunks, current, 0, line);
}

/// `pool.isShutdown()` / `pool.isTerminated()`.
pub fn emit_is_shutdown(chunks: &mut [Chunk], current: usize, line: u32) {
    let pool = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], pool, line);
    prop_get(chunks, current, pool, SHUTDOWN, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `pool.awaitTermination(t, unit)` — everything joined ⇒ true.
pub fn emit_await_termination(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let pool = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], pool, line);
    drain_tasks(chunks, current, pool, line);
    chunks[current].emit_bool_const(true, line);
}
