//! Dart core library adapters for Duration, DateTime, and Uri.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_compiler::primitives::{collections, reflection};

fn key(chunk: &mut Chunk, name: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(name)))
}

fn set_field(chunk: &mut Chunk, name: &str, line: u32) {
    let k = key(chunk, name);
    chunk.emit_op_u16(Op::STRUCT_SET, k, line);
    chunk.emit_op(Op::DROP, line);
}

fn get_field(chunk: &mut Chunk, name: &str, line: u32) {
    let k = key(chunk, name);
    chunk.emit_op_u16(Op::STRUCT_GET, k, line);
}

fn emit_slot_is_bigint(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "wasm:js-bigint", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_bigint_i32(chunks: &mut [Chunk], current: usize, value: i32, line: u32) {
    chunks[current].emit_i32_const(value, line);
    host::emit(&mut chunks[current], "ecma:bigint", "BigInt", 1, line);
}

fn set_string(chunk: &mut Chunk, name: &str, value: &str, line: u32) {
    core_wasm::dup(chunk, line);
    chunk.emit_string_const(value, line);
    set_field(chunk, name, line);
}

fn set_bool(chunk: &mut Chunk, name: &str, value: bool, line: u32) {
    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(value, line);
    set_field(chunk, name, line);
}

fn obj_new(chunk: &mut Chunk, line: u32) {
    host::emit(chunk, "ecma:object", "new", 0, line);
}

fn stamp_runtime_type(
    chunk: &mut Chunk,
    type_name: &str,
    kind: reflection::ReflectKind,
    line: u32,
) {
    set_string(chunk, reflection::FIELD_TYPE, type_name, line);
    set_string(chunk, reflection::FIELD_TYPE_NAME, type_name, line);
    set_string(chunk, reflection::FIELD_KIND, kind.as_str(), line);
}

fn date_get(chunks: &mut [Chunk], current: usize, ms_slot: u16, getter: &'static str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    host::emit(&mut chunks[current], "ecma:date", getter, 1, line);
}

fn wrap_duration_ms(chunk: &mut Chunk, line: u32) {
    let ms = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, ms, line);
    obj_new(chunk, line);
    stamp_runtime_type(chunk, "Duration", reflection::ReflectKind::Object, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    set_field(chunk, "inMilliseconds", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(1000.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    set_field(chunk, "inMicroseconds", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(1000.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inSeconds", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(60_000.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inMinutes", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(3_600_000.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inHours", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(86_400_000.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    set_field(chunk, "inDays", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(0.0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    set_field(chunk, "isNegative", line);
}

fn duration_ms_from_obj(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(chunk, "inMilliseconds", line);
}

pub fn emit_duration_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_f64_const(0.0, line);
    }
    wrap_duration_ms(&mut chunks[current], line);
}

pub fn emit_duration_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_f64_const(0.0, line);
    wrap_duration_ms(&mut chunks[current], line);
}

pub fn emit_duration_abs(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    duration_ms_from_obj(&mut chunks[current], slot, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    wrap_duration_ms(&mut chunks[current], line);
}

fn emit_slot_is_type(chunk: &mut Chunk, slot: u16, type_name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(chunk, reflection::FIELD_TYPE, line);
    chunk.emit_string_const(type_name, line);
    chunk.emit_op(Op::EQ, line);
}

pub fn emit_dart_abs(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    emit_slot_is_bigint(&mut chunks[current], value, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    crate::emitter::string_adapter::emit_dart_bigint_abs(chunks, current, line);
    chunks[current].emit_else(line);
    emit_slot_is_type(&mut chunks[current], value, "Duration", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_duration_abs(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_num_floor(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::F64_FLOOR, line);
}

pub fn emit_num_ceil(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::F64_CEIL, line);
}

pub fn emit_num_round(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:math", "round", 1, line);
}

pub fn emit_num_truncate(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::F64_TRUNC, line);
}

pub fn emit_num_to_double(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_num_remainder(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::math::emit_c_fmod(&mut chunks[current], line);
}

pub fn emit_num_is_negative(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    emit_slot_is_type(&mut chunks[current], value, "Duration", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    get_field(&mut chunks[current], "isNegative", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_f64_const(0.0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

pub fn emit_num_is_infinite(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "ecma:number", "isFinite", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "ecma:number", "isNaN", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);

    chunks[current].emit_op(Op::I32_AND, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_num_sign(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_duration_negate(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    duration_ms_from_obj(&mut chunks[current], slot, line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    wrap_duration_ms(&mut chunks[current], line);
}

/// Lower `DateTime(year[, month[, day[, hour[, minute[, second]]]]])` onto
/// `ecma:date.UTC`. Dart §`DateTime` gives every component after `year` a
/// default — `month` and `day` are **1**, the time parts are 0 — so only the
/// `argc` values actually on the stack may be popped; popping a fixed six (or
/// even a fixed three) would consume operands that were never pushed and read
/// whatever happened to be underneath (`DateTime(2000)` → undefined,
/// `DateTime(2000, 1)` → year 2066).
fn utc_from_stack(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    // Slots in ARGUMENT order, so index i is the i-th constructor parameter.
    let slots = [
        base,     // year
        base + 1, // month
        base + 2, // day
        base + 3, // hour
        base + 4, // minute
        base + 5, // second
    ];
    let supplied = (argc as usize).min(slots.len());
    // The last argument is on top, so pop right-to-left.
    for i in (0..supplied).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, slots[i], line);
    }
    // Fill in the omitted components with Dart's defaults.
    for (i, slot) in slots.iter().enumerate().skip(supplied) {
        // month/day are 1-based; hour/minute/second start at 0.
        let default = if i <= 2 { 1.0 } else { 0.0 };
        chunks[current].emit_f64_const(default, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    // `ecma:date.UTC` takes a 0-based month; Dart's is 1-based.
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    for slot in &slots[2..] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
    }
    host::emit(&mut chunks[current], "ecma:date", "UTC", 6, line);
}

fn wrap_datetime_ms(chunks: &mut [Chunk], current: usize, is_utc: bool, line: u32) {
    let ms = chunks[current].alloc_scratch(2);
    let dow = ms + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, ms, line);
    obj_new(&mut chunks[current], line);
    stamp_runtime_type(
        &mut chunks[current],
        "DateTime",
        reflection::ReflectKind::Object,
        line,
    );
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms, line);
    set_field(&mut chunks[current], "millisecondsSinceEpoch", line);
    for (field, getter, add_one) in [
        ("year", "getUTCFullYear", false),
        ("month", "getUTCMonth", true),
        ("day", "getUTCDate", false),
        ("hour", "getUTCHours", false),
        ("minute", "getUTCMinutes", false),
        ("second", "getUTCSeconds", false),
    ] {
        core_wasm::dup(&mut chunks[current], line);
        date_get(chunks, current, ms, getter, line);
        if add_one {
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_ADD, line);
        }
        set_field(&mut chunks[current], field, line);
    }
    date_get(chunks, current, ms, "getUTCDay", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dow, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dow, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dow, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(7.0, line);
    chunks[current].emit_end(line);
    set_field(&mut chunks[current], "weekday", line);
    set_bool(&mut chunks[current], "isUtc", is_utc, line);
}

fn datetime_ms_from_obj(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(chunk, "millisecondsSinceEpoch", line);
}

fn comparable_value_from_obj(chunk: &mut Chunk, slot: u16, out: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(chunk, "millisecondsSinceEpoch", line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(chunk, "inMilliseconds", line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_compare_slots(chunk: &mut Chunk, left: u16, right: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, left, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_datetime_new(chunks: &mut [Chunk], current: usize, argc: u8, is_utc: bool, line: u32) {
    utc_from_stack(chunks, current, argc, line);
    wrap_datetime_ms(chunks, current, is_utc, line);
}

pub fn emit_datetime_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let dur = chunks[current].alloc_scratch(2);
    let dt = dur + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, dur, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dt, line);
    datetime_ms_from_obj(&mut chunks[current], dt, line);
    duration_ms_from_obj(&mut chunks[current], dur, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    wrap_datetime_ms(chunks, current, false, line);
}

pub fn emit_dart_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(2);
    let receiver = value + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver, line);
    get_field(&mut chunks[current], reflection::FIELD_TYPE, line);
    chunks[current].emit_string_const("DateTime", line);
    chunks[current].emit_op(Op::EQ, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_datetime_add(chunks, current, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_push(chunks, current, line);

    chunks[current].emit_end(line);
}

pub fn emit_datetime_subtract(chunks: &mut [Chunk], current: usize, line: u32) {
    let dur = chunks[current].alloc_scratch(2);
    let dt = dur + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, dur, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dt, line);
    datetime_ms_from_obj(&mut chunks[current], dt, line);
    duration_ms_from_obj(&mut chunks[current], dur, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    wrap_datetime_ms(chunks, current, false, line);
}

pub fn emit_datetime_difference(chunks: &mut [Chunk], current: usize, line: u32) {
    let right = chunks[current].alloc_scratch(2);
    let left = right + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, right, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left, line);
    datetime_ms_from_obj(&mut chunks[current], left, line);
    datetime_ms_from_obj(&mut chunks[current], right, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    wrap_duration_ms(&mut chunks[current], line);
}

fn compare_ms(chunks: &mut [Chunk], current: usize, line: u32, op: Op) {
    let right = chunks[current].alloc_scratch(2);
    let left = right + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, right, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left, line);
    datetime_ms_from_obj(&mut chunks[current], left, line);
    datetime_ms_from_obj(&mut chunks[current], right, line);
    chunks[current].emit_op(op, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_datetime_is_before(chunks: &mut [Chunk], current: usize, line: u32) {
    compare_ms(chunks, current, line, Op::F64_LT);
}

pub fn emit_datetime_is_after(chunks: &mut [Chunk], current: usize, line: u32) {
    compare_ms(chunks, current, line, Op::F64_GT);
}

pub fn emit_datetime_same_moment(chunks: &mut [Chunk], current: usize, line: u32) {
    compare_ms(chunks, current, line, Op::F64_EQ);
}

pub fn emit_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let right = chunks[current].alloc_scratch(5);
    let left = right + 1;
    let r = right + 2;
    let l = right + 3;
    let method = right + 4;
    chunks[current].emit_op_u16(Op::LOCAL_SET, right, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left, line);
    emit_slot_is_bigint(&mut chunks[current], left, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    host::emit(&mut chunks[current], "ecma:bigint", "lt", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    host::emit(&mut chunks[current], "ecma:bigint", "gt", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    host::emit(&mut chunks[current], "wasm:js-string", "compare", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    host::emit(&mut chunks[current], "wasm:js-number", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_compare_slots(&mut chunks[current], left, right, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    get_field(&mut chunks[current], "compareTo", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, method, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    chunks[current].emit_else(line);
    comparable_value_from_obj(&mut chunks[current], left, l, line);
    comparable_value_from_obj(&mut chunks[current], right, r, line);
    emit_compare_slots(&mut chunks[current], l, r, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn get_url_prop(chunk: &mut Chunk, url_slot: u16, prop: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, url_slot, line);
    get_field(chunk, prop, line);
}

fn wrap_url(chunks: &mut [Chunk], current: usize, line: u32) {
    let url = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, url, line);
    obj_new(&mut chunks[current], line);
    stamp_runtime_type(
        &mut chunks[current],
        "Uri",
        reflection::ReflectKind::Object,
        line,
    );
    set_bool(&mut chunks[current], "__dart_uri_marker", true, line);
    core_wasm::dup(&mut chunks[current], line);
    get_url_prop(&mut chunks[current], url, "protocol", line);
    chunks[current].emit_string_const(":", line);
    chunks[current].emit_string_const("", line);
    host::emit(&mut chunks[current], "ecma:string", "replace", 3, line);
    set_field(&mut chunks[current], "scheme", line);
    for (dart, web) in [
        ("host", "hostname"),
        ("authority", "host"),
        ("path", "pathname"),
        ("queryParameters", "searchParams"),
        ("origin", "origin"),
    ] {
        core_wasm::dup(&mut chunks[current], line);
        get_url_prop(&mut chunks[current], url, web, line);
        set_field(&mut chunks[current], dart, line);
    }
    core_wasm::dup(&mut chunks[current], line);
    get_url_prop(&mut chunks[current], url, "port", line);
    set_field(&mut chunks[current], "port", line);
    core_wasm::dup(&mut chunks[current], line);
    get_url_prop(&mut chunks[current], url, "search", line);
    chunks[current].emit_string_const("?", line);
    chunks[current].emit_string_const("", line);
    host::emit(&mut chunks[current], "ecma:string", "replace", 3, line);
    set_field(&mut chunks[current], "query", line);
    core_wasm::dup(&mut chunks[current], line);
    get_url_prop(&mut chunks[current], url, "hash", line);
    chunks[current].emit_string_const("#", line);
    chunks[current].emit_string_const("", line);
    host::emit(&mut chunks[current], "ecma:string", "replace", 3, line);
    set_field(&mut chunks[current], "fragment", line);
    core_wasm::dup(&mut chunks[current], line);
    get_url_prop(&mut chunks[current], url, "href", line);
    set_field(&mut chunks[current], "href", line);
    set_bool(&mut chunks[current], "hasScheme", true, line);
    set_bool(&mut chunks[current], "hasAuthority", true, line);
    set_bool(&mut chunks[current], "isAbsolute", true, line);
}

pub fn emit_uri_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "web:url", "parse", 1, line);
    wrap_url(chunks, current, line);
}

pub fn emit_uri_http(chunks: &mut [Chunk], current: usize, argc: u8, https: bool, line: u32) {
    if argc > 2 {
        chunks[current].emit_op(Op::DROP, line);
    }
    if argc > 3 {
        chunks[current].emit_op(Op::DROP, line);
    }
    let path = chunks[current].alloc_scratch(2);
    let host_slot = path + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, host_slot, line);
    chunks[current].emit_string_const(if https { "https://" } else { "http://" }, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, host_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    emit_uri_parse(chunks, current, line);
}

pub fn emit_uri_file(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("file://", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    emit_uri_parse(chunks, current, line);
}

pub fn emit_uri_normalize_path(_chunks: &mut [Chunk], _current: usize, _line: u32) {}
pub fn emit_uri_replace(_chunks: &mut [Chunk], _current: usize, _line: u32) {}

pub fn emit_uri_resolve(chunks: &mut [Chunk], current: usize, line: u32) {
    let rel = chunks[current].alloc_scratch(2);
    let base = rel + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, rel, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rel, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    get_field(&mut chunks[current], "href", line);
    host::emit(&mut chunks[current], "web:url", "parse", 2, line);
    wrap_url(chunks, current, line);
}

pub fn emit_uri_resolve_uri(chunks: &mut [Chunk], current: usize, line: u32) {
    let rel = chunks[current].alloc_scratch(2);
    let base = rel + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, rel, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rel, line);
    get_field(&mut chunks[current], "href", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    get_field(&mut chunks[current], "href", line);
    host::emit(&mut chunks[current], "web:url", "parse", 2, line);
    wrap_url(chunks, current, line);
}

pub fn emit_list_filled(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(2);
    let length = value + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, length, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, length, line);
    collections::emit_new_with_length(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, length, line);
    collections::emit_fill(chunks, current, line);
}
