//! .NET `System.Diagnostics.Stopwatch` adapter — bytecode-only.
//!
//! Uses `wasi:clocks/monotonic-clock.now` (nanoseconds since process start)
//! for all timing.  The Stopwatch object layout:
//!   `{__type: "Stopwatch", __start_ns: f64, __accumulated_ns: f64, isrunning: bool}`

use crate::emitter::instructions::core_wasm;
use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn struct_get(chunk: &mut Chunk, field: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(field)));
    chunk.emit_op_u16(Op::STRUCT_GET, idx, line);
}

fn struct_set_drop(chunk: &mut Chunk, field: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(field)));
    chunk.emit_op_u16(Op::STRUCT_SET, idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_monotonic_now(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[0].add_import("wasi:clocks/monotonic-clock", "now");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(0u8, line);
}

/// Create a stopped Stopwatch object. Stack: `[]` → `[sw]`.
pub fn emit_stopwatch_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Stopwatch")), line);
    struct_set_drop(chunk, "__type", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_drop(chunk, "__start_ns", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_drop(chunk, "__accumulated_ns", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_drop(chunk, "isrunning", line);
}

/// `Stopwatch.StartNew()` — create and immediately start. Stack: `[]` → `[sw]`.
pub fn emit_stopwatch_start_new(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_stopwatch_new(chunks, current, line);
    // sw is on stack; DUP and start it
    core_wasm::dup(&mut chunks[current], line);
    emit_stopwatch_start_impl(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `sw.Start()` — start if not already running. Stack: `[sw]` → `[]`.
pub fn emit_stopwatch_start(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_stopwatch_start_impl(chunks, current, line);
}

fn emit_stopwatch_start_impl(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let sw_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, sw_slot, line);

    // if not isrunning: set __start_ns = now, set isrunning = true
    chunk.emit_op_u16(Op::LOCAL_GET, sw_slot, line);
    struct_get(chunk, "isrunning", line);
    chunk.emit_if(line);
    // already running — nothing to do
    chunk.emit_else(line);
    emit_monotonic_now(chunks, current, line);
    let chunk = &mut chunks[current];
    let now_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, now_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sw_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, now_slot, line);
    struct_set_drop(chunk, "__start_ns", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(true), line);
    struct_set_drop(chunk, "isrunning", line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
}

/// `sw.Stop()` — stop if running, accumulate elapsed. Stack: `[sw]` → `[]`.
pub fn emit_stopwatch_stop(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let sw_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, sw_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, sw_slot, line);
    struct_get(chunk, "isrunning", line);
    chunk.emit_if(line);
    // running — compute elapsed and accumulate
    emit_monotonic_now(chunks, current, line);
    let chunk = &mut chunks[current];
    let now_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, now_slot, line);

    // elapsed_ns = now - __start_ns
    chunk.emit_op_u16(Op::LOCAL_GET, now_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sw_slot, line);
    struct_get(chunk, "__start_ns", line);
    chunk.emit_op(Op::F64_SUB, line);
    let elapsed_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, elapsed_slot, line);

    // __accumulated_ns += elapsed_ns
    chunk.emit_op_u16(Op::LOCAL_GET, sw_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sw_slot, line);
    struct_get(chunk, "__accumulated_ns", line);
    chunk.emit_op_u16(Op::LOCAL_GET, elapsed_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    struct_set_drop(chunk, "__accumulated_ns", line);

    // isrunning = false
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_drop(chunk, "isrunning", line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
}

/// `sw.Reset()` — clear all timing state. Stack: `[sw]` → `[]`.
pub fn emit_stopwatch_reset(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_drop(chunk, "__start_ns", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_drop(chunk, "__accumulated_ns", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_drop(chunk, "isrunning", line);
    chunk.emit_op(Op::DROP, line);
}

/// `sw.Restart()` — reset then start. Stack: `[sw]` → `[]`.
pub fn emit_stopwatch_restart(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::dup(&mut chunks[current], line);
    emit_stopwatch_reset(chunks, current, line);
    core_wasm::dup(&mut chunks[current], line);
    emit_stopwatch_start_impl(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `sw.ElapsedMilliseconds` — total elapsed ms. Stack: `[sw]` → `[ms]`.
pub fn emit_stopwatch_elapsed_ms(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let sw_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, sw_slot, line);

    // if running: (now - __start_ns + __accumulated_ns) / 1e6
    // else: __accumulated_ns / 1e6
    chunk.emit_op_u16(Op::LOCAL_GET, sw_slot, line);
    struct_get(chunk, "isrunning", line);
    chunk.emit_if(line);
    emit_monotonic_now(chunks, current, line);
    let chunk = &mut chunks[current];
    let now_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, now_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, now_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sw_slot, line);
    struct_get(chunk, "__start_ns", line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sw_slot, line);
    struct_get(chunk, "__accumulated_ns", line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, sw_slot, line);
    struct_get(chunk, "__accumulated_ns", line);
    chunk.emit_end(line);
    // convert ns → ms
    chunk.emit_f64_const(1_000_000.0, line);
    chunk.emit_op(Op::F64_DIV, line);
}

/// `sw.IsRunning` — Stack: `[sw]` → `[bool]`.
pub fn emit_stopwatch_is_running(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let key = chunk.add_constant(Value::String(Arc::from("isrunning")));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
}
