//! .NET `System.DateTime` adapter — bytecode-only.
//!
//! `DateTime` is a calendar instant; .NET's `DateTime.Now /
//! UtcNow / Parse / Today` static methods plus the `New
//! DateTime(y, m, d, ...)` constructor produce a value with
//! `__type=DateTime` and a millisecond-since-epoch payload.
//!
//! The underlying primitive is `wasi:clocks/wall-clock.now` (WASI
//! 0.2.11 spec primitive — registered alongside the legacy flat
//! `wasi:clocks` namespace). `ecma:date.now` reads through it and
//! returns ms since epoch — the form ECMA-262 §21.4 [[DateValue]]
//! uses. Each adapter wraps that ms in a DateTime-shaped Object so
//! the .NET surface looks .NET-shaped while the bytecode is
//! standardized.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
use vybe_emitter::instructions::core_wasm;

use super::timespan_adapter;

const TYPE_KEY: &str = "__type";
const TIME_KEY: &str = "__time";

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

fn struct_set_named_field(chunk: &mut Chunk, key: &str, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, key_idx, line);
}

fn struct_set_named_field_drop(chunk: &mut Chunk, key: &str, line: u32) {
    struct_set_named_field(chunk, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn struct_get_named_field(chunk: &mut Chunk, key: &str, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_GET, key_idx, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    func: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, func);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_dt_getter(chunks: &mut [Chunk], current: usize, ms_slot: u16, getter: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    call_import(chunks, current, "ecma:date", getter, 1, line);
}

fn emit_datetime_time_from_obj(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from(TIME_KEY)), line);
    let idx = chunk.add_import("ecma:object", "get");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(2, line);
}

fn emit_named_field_from_obj(chunk: &mut Chunk, obj_slot: u16, field: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from(field)), line);
    let idx = chunk.add_import("ecma:object", "get");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(2, line);
}

fn emit_compare_numeric_slots(chunk: &mut Chunk, left_slot: u16, right_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::I32(-1), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_emitter::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_else(line);
    push_const(chunk, Value::I32(0), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_day_of_week_string(chunk: &mut Chunk, slot: u16, line: u32) {
    let done = chunk.emit_block(line);
    for (index, name) in [
        (0, "Sunday"),
        (1, "Monday"),
        (2, "Tuesday"),
        (3, "Wednesday"),
        (4, "Thursday"),
        (5, "Friday"),
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        push_const(chunk, Value::I32(index), line);
        vybe_emitter::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::String(Arc::from(name)), line);
        chunk.emit_br(1, line);
        chunk.emit_end(line);
    }
    push_const(chunk, Value::String(Arc::from("Saturday")), line);
    chunk.emit_end(line);
    chunk.patch_block(done);
}

fn emit_utc_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    year_slot: u16,
    month_slot: u16,
    day_slot: u16,
    hour_slot: Option<u16>,
    minute_slot: Option<u16>,
    second_slot: Option<u16>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, year_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, month_slot, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_GET, day_slot, line);
    if let Some(slot) = hour_slot {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        push_const(chunk, Value::I32(0), line);
    }
    if let Some(slot) = minute_slot {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        push_const(chunk, Value::I32(0), line);
    }
    if let Some(slot) = second_slot {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        push_const(chunk, Value::I32(0), line);
    }
    call_import(chunks, current, "ecma:date", "UTC", 6, line);
}

fn emit_wrap_ms_internal(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
    include_composites: bool,
) {
    let chunk = &mut chunks[current];
    let ms_slot = chunk.alloc_scratch(9);
    let year_slot = ms_slot + 1;
    let month_slot = ms_slot + 2;
    let day_slot = ms_slot + 3;
    let hour_slot = ms_slot + 4;
    let minute_slot = ms_slot + 5;
    let second_slot = ms_slot + 6;
    let dow_slot = ms_slot + 7;
    let obj_slot = ms_slot + 8;

    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);

    emit_dt_getter(chunks, current, ms_slot, "getUTCFullYear", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, year_slot, line);

    emit_dt_getter(chunks, current, ms_slot, "getUTCMonth", line);
    push_const(&mut chunks[current], Value::I32(1), line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, month_slot, line);

    emit_dt_getter(chunks, current, ms_slot, "getUTCDate", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, day_slot, line);

    emit_dt_getter(chunks, current, ms_slot, "getUTCHours", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hour_slot, line);

    emit_dt_getter(chunks, current, ms_slot, "getUTCMinutes", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, minute_slot, line);

    emit_dt_getter(chunks, current, ms_slot, "getUTCSeconds", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, second_slot, line);

    emit_dt_getter(chunks, current, ms_slot, "getUTCDay", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dow_slot, line);

    let chunk = &mut chunks[current];
    let object_new = chunk.add_import("ecma:object", "new");
    chunk.emit_op_u16(Op::CALL_IMPORT, object_new, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("DateTime")), line);
    struct_set_named_field_drop(chunk, TYPE_KEY, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    struct_set_named_field_drop(chunk, TIME_KEY, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("Unspecified")), line);
    struct_set_named_field_drop(chunk, "Kind", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(10_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    push_const(chunk, Value::F64(621_355_968_000_000_000.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    struct_set_named_field_drop(chunk, "Ticks", line);

    for (field, slot) in [
        ("Year", year_slot),
        ("Month", month_slot),
        ("Day", day_slot),
        ("Hour", hour_slot),
        ("Minute", minute_slot),
        ("Second", second_slot),
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        struct_set_named_field_drop(chunk, field, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    emit_day_of_week_string(chunk, dow_slot, line);
    struct_set_named_field_drop(chunk, "DayOfWeek", line);

    if include_composites {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        emit_utc_from_slots(
            chunks, current, year_slot, month_slot, day_slot, None, None, None, line,
        );
        emit_wrap_ms_internal(chunks, current, line, false);
        struct_set_named_field_drop(&mut chunks[current], "Date", line);

        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, hour_slot, line);
        push_const(chunk, Value::F64(3_600_000.0), line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op_u16(Op::LOCAL_GET, minute_slot, line);
        push_const(chunk, Value::F64(60_000.0), line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_GET, second_slot, line);
        push_const(chunk, Value::F64(1000.0), line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        timespan_adapter::emit_build_timespan_from_total_ms(chunk, line);
        struct_set_named_field_drop(chunk, "TimeOfDay", line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// Wrap a millisecond timestamp on stack-top as a DateTime object.
/// Stack on entry: `[ms]` ; Stack on exit: `[datetime_obj]`.
fn emit_wrap_ms(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_wrap_ms_internal(chunks, current, line, true);
}

/// Wrap a millisecond timestamp on stack-top as a DateTime object for
/// other .NET adapters.
pub fn emit_datetime_from_millis(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_wrap_ms(chunks, current, line);
}

/// `DateTime.Now` / `DateTime.UtcNow` — read `ecma:date.now` (which
/// reads through `wasi:clocks/wall-clock.now`) and wrap in a
/// DateTime object.
///
/// Stack on entry: `[]` ; Stack on exit: `[datetime_obj]`
pub fn emit_datetime_now(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:date", "now", 0, line);
    emit_wrap_ms(chunks, current, line);
}

/// `DateTime.Parse(s)` — parse a date string via `ecma:date.parse`
/// (ECMA-262 §21.4.3.2) and wrap.
///
/// Stack on entry: `[s]` ; Stack on exit: `[datetime_obj]`
pub fn emit_datetime_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:date", "parse", 1, line);
    emit_wrap_ms(chunks, current, line);
}

/// `DateTime.Today` — synonym for `Now` in .NET (returns midnight of
/// today; we return the current instant for the MVP). Same bytecode
/// as `Now`.
pub fn emit_datetime_today(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:date", "now", 0, line);
    let chunk = &mut chunks[current];
    let ms_slot = chunk.alloc_scratch(4);
    let year_slot = ms_slot + 1;
    let month_slot = ms_slot + 2;
    let day_slot = ms_slot + 3;
    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    emit_dt_getter(chunks, current, ms_slot, "getUTCFullYear", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, year_slot, line);
    emit_dt_getter(chunks, current, ms_slot, "getUTCMonth", line);
    push_const(&mut chunks[current], Value::I32(1), line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, month_slot, line);
    emit_dt_getter(chunks, current, ms_slot, "getUTCDate", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, day_slot, line);
    emit_utc_from_slots(
        chunks, current, year_slot, month_slot, day_slot, None, None, None, line,
    );
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    match argc {
        3 | 6 | 7 => {
            let second_slot = chunk.alloc_scratch(7);
            let minute_slot = second_slot + 1;
            let hour_slot = second_slot + 2;
            let day_slot = second_slot + 3;
            let month_slot = second_slot + 4;
            let year_slot = second_slot + 5;
            let kind_slot = second_slot + 6;

            if argc == 7 {
                chunk.emit_op_u16(Op::LOCAL_SET, kind_slot, line);
            }
            if argc >= 6 {
                chunk.emit_op_u16(Op::LOCAL_SET, second_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, minute_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, hour_slot, line);
            }
            chunk.emit_op_u16(Op::LOCAL_SET, day_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, month_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, year_slot, line);

            emit_utc_from_slots(
                chunks,
                current,
                year_slot,
                month_slot,
                day_slot,
                if argc >= 6 { Some(hour_slot) } else { None },
                if argc >= 6 { Some(minute_slot) } else { None },
                if argc >= 6 { Some(second_slot) } else { None },
                line,
            );
            emit_wrap_ms(chunks, current, line);
            if argc == 7 {
                let chunk = &mut chunks[current];
                core_wasm::dup(chunk, line);
                chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
                struct_set_named_field(chunk, "Kind", line);
            }
        }
        _ => emit_datetime_now(chunks, current, line),
    }
}

pub fn emit_datetime_year(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "Year", line);
}

pub fn emit_datetime_month(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "Month", line);
}

pub fn emit_datetime_day(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "Day", line);
}

pub fn emit_datetime_hour(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "Hour", line);
}

pub fn emit_datetime_minute(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "Minute", line);
}

pub fn emit_datetime_second(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "Second", line);
}

pub fn emit_datetime_day_of_week(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "DayOfWeek", line);
}

pub fn emit_datetime_add_days(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.alloc_scratch(2);
    let date_slot = value_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
    emit_datetime_time_from_obj(chunk, date_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    push_const(chunk, Value::F64(86_400_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_add_hours(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.alloc_scratch(2);
    let date_slot = value_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
    emit_datetime_time_from_obj(chunk, date_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    push_const(chunk, Value::F64(3_600_000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_add_months(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let months_slot = chunk.alloc_scratch(9);
    let date_slot = months_slot + 1;
    let year_slot = months_slot + 2;
    let month_slot = months_slot + 3;
    let day_slot = months_slot + 4;
    let hour_slot = months_slot + 5;
    let minute_slot = months_slot + 6;
    let second_slot = months_slot + 7;
    let total_months_slot = months_slot + 8;
    chunk.emit_op_u16(Op::LOCAL_SET, months_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);

    emit_named_field_from_obj(chunk, date_slot, "Year", line);
    chunk.emit_op_u16(Op::LOCAL_SET, year_slot, line);
    emit_named_field_from_obj(chunk, date_slot, "Month", line);
    chunk.emit_op_u16(Op::LOCAL_SET, month_slot, line);
    emit_named_field_from_obj(chunk, date_slot, "Day", line);
    chunk.emit_op_u16(Op::LOCAL_SET, day_slot, line);
    emit_named_field_from_obj(chunk, date_slot, "Hour", line);
    chunk.emit_op_u16(Op::LOCAL_SET, hour_slot, line);
    emit_named_field_from_obj(chunk, date_slot, "Minute", line);
    chunk.emit_op_u16(Op::LOCAL_SET, minute_slot, line);
    emit_named_field_from_obj(chunk, date_slot, "Second", line);
    chunk.emit_op_u16(Op::LOCAL_SET, second_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, year_slot, line);
    push_const(chunk, Value::F64(12.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, month_slot, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, months_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, total_months_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, total_months_slot, line);
    push_const(chunk, Value::F64(12.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_TRUNC, line);
    chunk.emit_op_u16(Op::LOCAL_SET, year_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, total_months_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, year_slot, line);
    push_const(chunk, Value::F64(12.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, month_slot, line);

    emit_utc_from_slots(
        chunks,
        current,
        year_slot,
        month_slot,
        day_slot,
        Some(hour_slot),
        Some(minute_slot),
        Some(second_slot),
        line,
    );
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_days_in_month(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let month_slot = chunk.alloc_scratch(3);
    let year_slot = month_slot + 1;
    let ms_slot = month_slot + 2;
    chunk.emit_op_u16(Op::LOCAL_SET, month_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, year_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, year_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, month_slot, line);
    push_const(chunk, Value::I32(1), line);
    push_const(chunk, Value::I32(0), line);
    push_const(chunk, Value::I32(0), line);
    push_const(chunk, Value::I32(0), line);
    call_import(chunks, current, "ecma:date", "UTC", 6, line);
    push_const(&mut chunks[current], Value::F64(86_400_000.0), line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    emit_dt_getter(chunks, current, ms_slot, "getUTCDate", line);
}

pub fn emit_datetime_is_leap_year(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let year_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, year_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, year_slot, line);
    push_const(chunk, Value::I32(2), line);
    emit_datetime_days_in_month(chunks, current, line);
    push_const(&mut chunks[current], Value::I32(29), line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
}

pub fn emit_datetime_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(4);
    let left_slot = right_slot + 1;
    let right_time_slot = right_slot + 2;
    let left_time_slot = right_slot + 3;
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    emit_datetime_time_from_obj(chunk, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_time_slot, line);
    emit_datetime_time_from_obj(chunk, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_time_slot, line);
    emit_compare_numeric_slots(chunk, left_time_slot, right_time_slot, line);
}

pub fn emit_datetime_to_short_date_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_datetime_time_from_obj(chunk, obj_slot, line);
    call_import(chunks, current, "ecma:date", "toISOString", 1, line);
}

pub fn emit_datetime_add_timespan(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let span_slot = chunk.alloc_scratch(2);
    let date_slot = span_slot + 1;
    let total_ms_key = string_key(chunk, "TotalMilliseconds");
    chunk.emit_op_u16(Op::LOCAL_SET, span_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
    emit_datetime_time_from_obj(chunk, date_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, span_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, total_ms_key, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_subtract_datetime(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(2);
    let left_slot = right_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    emit_datetime_time_from_obj(chunk, left_slot, line);
    emit_datetime_time_from_obj(chunk, right_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    timespan_adapter::emit_build_timespan_from_total_ms(chunk, line);
}
