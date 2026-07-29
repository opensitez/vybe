//! .NET `System.TimeSpan` adapter — bytecode-only.
//!
//! `TimeSpan` is a duration value type; .NET's
//! `TimeSpan.From{Days,Hours,Minutes,Seconds,Milliseconds}(n)`
//! factory methods build a duration record from a unit count.
//! There's no ECMA-262 mirror (JS uses raw `number` ms), but the
//! arithmetic is trivial: multiply by the unit-to-ms factor and
//! stash on a struct.
//!
//! Each adapter emits inline bytecode — no host fns. The result has
//! shape `{ __type: "TimeSpan", totalmilliseconds, totalseconds,
//! totalminutes, totalhours, totaldays, days, hours, minutes,
//! seconds }` matching the existing `vybe:types/timeSpan*` host
//! impls so callers continue to work.

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::instructions::{core_wasm, host};

use vybe_compiler::primitives::math;

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn string_key(chunk: &mut Chunk, key: &str) -> u16 {
    if let Some((idx, _)) = chunk
        .constants
        .iter()
        .enumerate()
        .find(|(_, value)| matches!(value, Value::String(s) if s.as_ref() == key))
    {
        idx as u16
    } else {
        chunk.add_constant(Value::String(Arc::from(key)))
    }
}

fn struct_set_field(chunk: &mut Chunk, key_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::STRUCT_SET, key_idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn struct_set_named_field(chunk: &mut Chunk, key: &str, line: u32) {
    let value_slot = chunk.alloc_scratch(2);
    let obj_slot = value_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from(key)), line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let idx = chunk.add_import("ecma:object", "set");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(3, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_array_get_const_index(chunk: &mut Chunk, array_slot: u16, index: f64, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    push_const(chunk, Value::F64(index), line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

fn emit_parse_number_from_slot(chunks: &mut [Chunk], current: usize, text_slot: u16, line: u32) {
    let parse_int_idx = chunks[current].add_import("ecma:number", "parseInt");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, parse_int_idx, line);
    chunk.emit(1, line);
    chunk.emit_op(Op::F64_FLOOR, line);
}

fn emit_store_array_part_as_number(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    index: f64,
    out_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    emit_array_get_const_index(chunk, array_slot, index, line);
    let text_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    emit_parse_number_from_slot(chunks, current, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
}

fn emit_total_ms_from_obj(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("TotalMilliseconds")), line);
    let idx = chunk.add_import("ecma:object", "get");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(2, line);
}

/// Build the TimeSpan object given the total milliseconds on the stack.
/// Stack on entry: `[total_ms]` ; Stack on exit: `[ts_obj]`
pub(crate) fn emit_build_timespan_from_total_ms(chunk: &mut Chunk, line: u32) {
    let ms_slot = chunk.alloc_scratch(6);
    let days_slot = ms_slot + 1;
    let rem_slot = ms_slot + 2;
    let hours_slot = ms_slot + 3;
    let minutes_slot = ms_slot + 4;
    let seconds_slot = ms_slot + 5;
    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);

    let type_key = string_key(chunk, "__type");
    let total_ms_key = string_key(chunk, "totalmilliseconds");
    let total_sec_key = string_key(chunk, "totalseconds");
    let total_min_key = string_key(chunk, "totalminutes");
    let total_hr_key = string_key(chunk, "totalhours");
    let total_day_key = string_key(chunk, "totaldays");
    let ticks_key = string_key(chunk, "ticks");

    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(86_400_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    math::emit_trunc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, days_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, days_slot, line);
    push_const(chunk, Value::F64(86_400_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rem_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    push_const(chunk, Value::F64(3_600_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    math::emit_trunc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hours_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, hours_slot, line);
    push_const(chunk, Value::F64(3_600_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rem_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    push_const(chunk, Value::F64(60_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    math::emit_trunc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, minutes_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, minutes_slot, line);
    push_const(chunk, Value::F64(60_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rem_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    math::emit_trunc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, seconds_slot, line);

    let object_new = chunk.add_import("ecma:object", "new");
    chunk.emit_op_u16(Op::CALL_IMPORT, object_new, line);
    chunk.emit(0, line);

    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("TimeSpan")), line);
    struct_set_field(chunk, type_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    struct_set_field(chunk, total_ms_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    struct_set_named_field(chunk, "TotalMilliseconds", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(10_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    struct_set_field(chunk, ticks_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(10_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    struct_set_named_field(chunk, "Ticks", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, total_sec_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_named_field(chunk, "TotalSeconds", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(60_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, total_min_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(60_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_named_field(chunk, "TotalMinutes", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(3_600_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, total_hr_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(3_600_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_named_field(chunk, "TotalHours", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(86_400_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_field(chunk, total_day_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(86_400_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    struct_set_named_field(chunk, "TotalDays", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, days_slot, line);
    struct_set_named_field(chunk, "Days", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, hours_slot, line);
    struct_set_named_field(chunk, "Hours", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, minutes_slot, line);
    struct_set_named_field(chunk, "Minutes", line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, seconds_slot, line);
    struct_set_named_field(chunk, "Seconds", line);
}

/// Build a TimeSpan from a count of `unit_ms` units. Stack: `[n]` →
/// `[ts]`. Internally: `total_ms = n * unit_ms`, then build the
/// record. Generic over unit so all `From*` methods share one body.
fn emit_timespan_from_unit(chunks: &mut [Chunk], current: usize, unit_ms: f64, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(unit_ms), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_build_timespan_from_total_ms(chunk, line);
}

/// `TimeSpan.FromDays(n)` — `n * 86_400_000` ms.
pub fn emit_timespan_from_days(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 86_400_000.0, line);
}

/// `TimeSpan.FromHours(n)` — `n * 3_600_000` ms.
pub fn emit_timespan_from_hours(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 3_600_000.0, line);
}

/// `TimeSpan.FromMinutes(n)` — `n * 60_000` ms.
pub fn emit_timespan_from_minutes(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 60_000.0, line);
}

/// `TimeSpan.FromSeconds(n)` — `n * 1000` ms.
pub fn emit_timespan_from_seconds(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 1000.0, line);
}

/// `TimeSpan.FromMilliseconds(n)` — pass-through.
pub fn emit_timespan_from_milliseconds(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_timespan_from_unit(chunks, current, 1.0, line);
}

/// `TimeSpan.Zero` — 0-duration TimeSpan. Stack: `[]` → `[ts]`.
pub fn emit_timespan_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(0.0), line);
    emit_build_timespan_from_total_ms(chunk, line);
}

pub fn emit_timespan_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let to_str_idx = chunks[current].add_import("ecma:string", "String");
    let chunk = &mut chunks[current];
    let text_slot = chunk.alloc_scratch(5);
    let parts_slot = text_slot + 1;
    let hours_slot = text_slot + 2;
    let minutes_slot = text_slot + 3;
    let seconds_slot = text_slot + 4;

    chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::String(Arc::from(":")), line);
    host::emit(chunk, "ecma:string", "split", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, parts_slot, line);

    emit_store_array_part_as_number(chunks, current, parts_slot, 0.0, hours_slot, line);
    emit_store_array_part_as_number(chunks, current, parts_slot, 1.0, minutes_slot, line);
    emit_store_array_part_as_number(chunks, current, parts_slot, 2.0, seconds_slot, line);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, hours_slot, line);
    push_const(chunk, Value::F64(3600.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, minutes_slot, line);
    push_const(chunk, Value::F64(60.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, seconds_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_build_timespan_from_total_ms(chunk, line);
}

fn emit_compare_numeric_slots(chunk: &mut Chunk, left_slot: u16, right_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::I32(-1), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_else(line);
    push_const(chunk, Value::I32(0), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_timespan_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    match argc {
        3 => {
            let seconds_slot = chunk.alloc_scratch(3);
            let minutes_slot = seconds_slot + 1;
            let hours_slot = seconds_slot + 2;

            chunk.emit_op_u16(Op::LOCAL_SET, seconds_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, minutes_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, hours_slot, line);

            chunk.emit_op_u16(Op::LOCAL_GET, hours_slot, line);
            push_const(chunk, Value::F64(3600.0), line);
            chunk.emit_op(Op::F64_MUL, line);
            chunk.emit_op_u16(Op::LOCAL_GET, minutes_slot, line);
            push_const(chunk, Value::F64(60.0), line);
            chunk.emit_op(Op::F64_MUL, line);
            chunk.emit_op(Op::F64_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_GET, seconds_slot, line);
            chunk.emit_op(Op::F64_ADD, line);
            push_const(chunk, Value::F64(1000.0), line);
            chunk.emit_op(Op::F64_MUL, line);
            emit_build_timespan_from_total_ms(chunk, line);
        }
        _ => emit_timespan_zero(chunks, current, line),
    }
}

pub fn emit_timespan_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(4);
    let left_slot = right_slot + 1;
    let right_ms_slot = right_slot + 2;
    let left_ms_slot = right_slot + 3;

    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);

    emit_total_ms_from_obj(chunk, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_ms_slot, line);
    emit_total_ms_from_obj(chunk, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_ms_slot, line);
    emit_compare_numeric_slots(chunk, left_ms_slot, right_ms_slot, line);
}

pub fn emit_timespan_negate(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_total_ms_from_obj(chunk, obj_slot, line);
    push_const(chunk, Value::F64(-1.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_build_timespan_from_total_ms(chunk, line);
}

pub fn emit_timespan_duration(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_total_ms_from_obj(chunk, obj_slot, line);
    math::emit_abs(chunk, line);
    emit_build_timespan_from_total_ms(chunk, line);
}

pub fn emit_timespan_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(2);
    let left_slot = right_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    emit_total_ms_from_obj(chunk, left_slot, line);
    emit_total_ms_from_obj(chunk, right_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_build_timespan_from_total_ms(chunk, line);
}

pub fn emit_timespan_sub(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(2);
    let left_slot = right_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    emit_total_ms_from_obj(chunk, left_slot, line);
    emit_total_ms_from_obj(chunk, right_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    emit_build_timespan_from_total_ms(chunk, line);
}
