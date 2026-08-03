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
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::object::{emit_bind_method, emit_bind_method_with_slot};

use super::timespan_adapter;

const TYPE_KEY: &str = "__type";
const TIME_KEY: &str = "__time";

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val) }
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
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
}

fn struct_set_named_field_drop(chunk: &mut Chunk, key: &str, line: u32) {
    struct_set_named_field(chunk, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn struct_get_named_field(chunk: &mut Chunk, key: &str, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key_idx, line);
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

fn create_function_chunk(name: &str, arity: u8) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = arity;
    c
}

fn bind_value_to_string(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, line: u32) {
    let mut method = create_function_chunk("__dotnet_value_tostring", 1);
    let value_key = method.add_constant(Value::String(Arc::from("Value")));
    method.emit_op_u16(Op::LOCAL_GET, 0, line);
    method.emit_struct_field_op(Op::STRUCT_GET, 0, value_key, line);
    method.emit_op(Op::RETURN, line);
    method.local_count = 1;
    chunks.push(method);
    let method_idx = chunks.len() - 1;

    emit_bind_method_with_slot(
        &mut chunks[current],
        obj_slot,
        "tostring",
        Some(vybe_ast::ProtocolSlot::ToString),
        method_idx,
        None,
        line,
    );
}

fn emit_dt_getter(chunks: &mut [Chunk], current: usize, ms_slot: u16, getter: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    call_import(chunks, current, "ecma:date", getter, 1, line);
}

fn emit_datetime_time_from_obj(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = string_key(chunk, TIME_KEY);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

fn emit_named_field_from_obj(chunk: &mut Chunk, obj_slot: u16, field: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = string_key(chunk, field);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
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

fn emit_i32_condition_as_bool(chunk: &mut Chunk, line: u32) {
    chunk.emit_if(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

fn emit_day_of_week_object(chunks: &mut Vec<Chunk>, current: usize, dow_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    let text_slot = chunk.alloc_scratch(2);
    let obj_slot = text_slot + 1;
    push_const(chunk, Value::String(Arc::from("Saturday")), line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    for (index, name) in [
        (0, "Sunday"),
        (1, "Monday"),
        (2, "Tuesday"),
        (3, "Wednesday"),
        (4, "Thursday"),
        (5, "Friday"),
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, dow_slot, line);
        push_const(chunk, Value::F64(f64::from(index)), line);
        chunk.emit_op(Op::F64_EQ, line);
        chunk.emit_if(line);
        push_const(chunk, Value::String(Arc::from(name)), line);
        chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
        chunk.emit_end(line);
    }
    call_import(chunks, current, "ecma:object", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("DayOfWeek")), line);
    struct_set_named_field_drop(chunk, TYPE_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    struct_set_named_field_drop(chunk, "Value", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    struct_set_named_field_drop(chunk, "value", line);
    bind_value_to_string(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
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
    millisecond_slot: Option<u16>,
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
    if let Some(slot) = millisecond_slot {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        call_import(chunks, current, "ecma:date", "UTC", 7, line);
    } else {
        call_import(chunks, current, "ecma:date", "UTC", 6, line);
    }
}

fn emit_wrap_ms_internal(
    chunks: &mut Vec<Chunk>,
    current: usize,
    line: u32,
    include_composites: bool,
) {
    let ms_slot = {
        let chunk = &mut chunks[current];
        let ms_slot = chunk.alloc_scratch(12);
        chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
        ms_slot
    };
    let year_slot = ms_slot + 1;
    let month_slot = ms_slot + 2;
    let day_slot = ms_slot + 3;
    let hour_slot = ms_slot + 4;
    let minute_slot = ms_slot + 5;
    let second_slot = ms_slot + 6;
    let dow_slot = ms_slot + 7;
    let millis_slot = ms_slot + 8;
    let year_start_slot = ms_slot + 9;
    let one_slot = ms_slot + 10;
    let obj_slot = ms_slot + 11;

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

    emit_dt_getter(chunks, current, ms_slot, "getUTCMilliseconds", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, millis_slot, line);

    emit_dt_getter(chunks, current, ms_slot, "getUTCDay", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dow_slot, line);

    push_const(&mut chunks[current], Value::I32(1), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, one_slot, line);
    emit_utc_from_slots(
        chunks,
        current,
        year_slot,
        one_slot,
        one_slot,
        None,
        None,
        None,
        None,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, year_start_slot, line);

    call_import(chunks, current, "ecma:object", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("datetime")), line);
    struct_set_named_field_drop(chunk, TYPE_KEY, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    struct_set_named_field_drop(chunk, TIME_KEY, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("Unspecified")), line);
    struct_set_named_field_drop(chunk, "Kind", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("Unspecified")), line);
    struct_set_named_field_drop(chunk, "kind", line);

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
        ("Millisecond", millis_slot),
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        struct_set_named_field_drop(chunk, field, line);
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        struct_set_named_field_drop(chunk, &field.to_ascii_lowercase(), line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, year_start_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    push_const(chunk, Value::F64(vybe_compiler::primitives::datetime::MS_PER_DAY), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_TRUNC, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_ADD, line);
    struct_set_named_field_drop(chunk, "DayOfYear", line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    emit_day_of_week_object(chunks, current, dow_slot, line);
    struct_set_named_field_drop(&mut chunks[current], "DayOfWeek", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    emit_day_of_week_object(chunks, current, dow_slot, line);
    struct_set_named_field_drop(&mut chunks[current], "dayofweek", line);

    if include_composites {
        chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        emit_utc_from_slots(
            chunks, current, year_slot, month_slot, day_slot, None, None, None, None, line,
        );
        emit_wrap_ms_internal(chunks, current, line, false);
        struct_set_named_field_drop(&mut chunks[current], "Date", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        emit_utc_from_slots(
            chunks, current, year_slot, month_slot, day_slot, None, None, None, None, line,
        );
        emit_wrap_ms_internal(chunks, current, line, false);
        struct_set_named_field_drop(&mut chunks[current], "date", line);

        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, hour_slot, line);
        push_const(chunk, Value::F64(vybe_compiler::primitives::datetime::MS_PER_HOUR), line);
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
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, hour_slot, line);
        push_const(chunk, Value::F64(vybe_compiler::primitives::datetime::MS_PER_HOUR), line);
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
        struct_set_named_field_drop(chunk, "timeofday", line);
    }

    bind_datetime_to_string(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// Wrap a millisecond timestamp on stack-top as a DateTime object.
/// Stack on entry: `[ms]` ; Stack on exit: `[datetime_obj]`.
fn emit_wrap_ms(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_wrap_ms_internal(chunks, current, line, true);
}

/// Wrap a millisecond timestamp on stack-top as a DateTime object for
/// other .NET adapters.
pub fn emit_datetime_from_millis(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_wrap_ms(chunks, current, line);
}

/// `DateTime.Now` / `DateTime.UtcNow` — read `ecma:date.now` (which
/// reads through `wasi:clocks/wall-clock.now`) and wrap in a
/// DateTime object.
///
/// Stack on entry: `[]` ; Stack on exit: `[datetime_obj]`
pub fn emit_datetime_now(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    call_import(chunks, current, "ecma:date", "now", 0, line);
    emit_wrap_ms(chunks, current, line);
}

/// `DateTime.Parse(s)` — parse a date string via `ecma:date.parse`
/// (ECMA-262 §21.4.3.2) and wrap.
///
/// Stack on entry: `[s]` ; Stack on exit: `[datetime_obj]`
pub fn emit_datetime_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    call_import(chunks, current, "ecma:date", "parse", 1, line);
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_try_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    call_import(chunks, current, "ecma:date", "parse", 1, line);
    let chunk = &mut chunks[current];
    let ms_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    emit_wrap_ms(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `DateTime.Today` — synonym for `Now` in .NET (returns midnight of
/// today; we return the current instant for the MVP). Same bytecode
/// as `Now`.
pub fn emit_datetime_today(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
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
        chunks, current, year_slot, month_slot, day_slot, None, None, None, None, line,
    );
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_min_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(-62_135_596_800_000.0), line);
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_max_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(253_402_300_799_999.0), line);
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    match argc {
        1 => {
            let ticks_slot = chunk.alloc_scratch(1);
            chunk.emit_op_u16(Op::LOCAL_SET, ticks_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, ticks_slot, line);
            push_const(chunk, Value::F64(621_355_968_000_000_000.0), line);
            chunk.emit_op(Op::F64_SUB, line);
            push_const(chunk, Value::F64(10_000.0), line);
            chunk.emit_op(Op::F64_DIV, line);
            emit_wrap_ms(chunks, current, line);
        }
        3 | 6 | 7 | 8 => {
            let second_slot = chunk.alloc_scratch(9);
            let minute_slot = second_slot + 1;
            let hour_slot = second_slot + 2;
            let day_slot = second_slot + 3;
            let month_slot = second_slot + 4;
            let year_slot = second_slot + 5;
            let millis_or_kind_slot = second_slot + 6;
            let kind_slot = second_slot + 7;
            let is_kind_slot = second_slot + 8;

            if argc == 8 {
                chunk.emit_op_u16(Op::LOCAL_SET, kind_slot, line);
            }
            if argc >= 7 {
                chunk.emit_op_u16(Op::LOCAL_SET, millis_or_kind_slot, line);
                if argc == 7 {
                    chunk.emit_op_u16(Op::LOCAL_GET, millis_or_kind_slot, line);
                    chunk.emit_op_u16(Op::LOCAL_SET, kind_slot, line);
                }
            }
            if argc >= 6 {
                chunk.emit_op_u16(Op::LOCAL_SET, second_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, minute_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, hour_slot, line);
            }
            chunk.emit_op_u16(Op::LOCAL_SET, day_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, month_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, year_slot, line);

            if argc == 7 {
                let is_string = chunk.add_import("wasm:js-string", "test");
                chunk.emit_op_u16(Op::LOCAL_GET, millis_or_kind_slot, line);
                chunk.emit_call(is_string, 1, line);
                chunk.emit_op_u16(Op::LOCAL_SET, is_kind_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, is_kind_slot, line);
                chunk.emit_if(line);
                push_const(chunk, Value::I32(0), line);
                chunk.emit_op_u16(Op::LOCAL_SET, millis_or_kind_slot, line);
                chunk.emit_end(line);
            }

            emit_utc_from_slots(
                chunks,
                current,
                year_slot,
                month_slot,
                day_slot,
                if argc >= 6 { Some(hour_slot) } else { None },
                if argc >= 6 { Some(minute_slot) } else { None },
                if argc >= 6 { Some(second_slot) } else { None },
                if argc >= 7 { Some(millis_or_kind_slot) } else { None },
                line,
            );
            emit_wrap_ms(chunks, current, line);
            if argc == 7 {
                let chunk = &mut chunks[current];
                chunk.emit_op_u16(Op::LOCAL_GET, is_kind_slot, line);
                chunk.emit_if(line);
                core_wasm::dup(chunk, line);
                chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
                struct_set_named_field_drop(chunk, "Kind", line);
                core_wasm::dup(chunk, line);
                chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
                struct_set_named_field_drop(chunk, "kind", line);
                chunk.emit_end(line);
            } else if argc == 8 {
                let chunk = &mut chunks[current];
                core_wasm::dup(chunk, line);
                chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
                struct_set_named_field_drop(chunk, "Kind", line);
                core_wasm::dup(chunk, line);
                chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
                struct_set_named_field_drop(chunk, "kind", line);
            }
        }
        _ => emit_datetime_now(chunks, current, line) }
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

pub fn emit_datetime_millisecond(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "Millisecond", line);
}

pub fn emit_datetime_day_of_year(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "DayOfYear", line);
}

pub fn emit_datetime_day_of_week(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "DayOfWeek", line);
}

pub fn emit_datetime_ticks(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "Ticks", line);
}

pub fn emit_datetime_kind(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "Kind", line);
}

pub fn emit_datetime_date(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "Date", line);
}

pub fn emit_datetime_time_of_day(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], "TimeOfDay", line);
}

pub fn emit_datetime_add_days(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.alloc_scratch(2);
    let date_slot = value_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
    emit_datetime_time_from_obj(chunk, date_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    push_const(chunk, Value::F64(vybe_compiler::primitives::datetime::MS_PER_DAY), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_add_ticks(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let ticks_slot = chunk.alloc_scratch(2);
    let date_slot = ticks_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, ticks_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
    emit_datetime_time_from_obj(chunk, date_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ticks_slot, line);
    push_const(chunk, Value::F64(10_000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_add_hours(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.alloc_scratch(2);
    let date_slot = value_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
    emit_datetime_time_from_obj(chunk, date_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    push_const(chunk, Value::F64(vybe_compiler::primitives::datetime::MS_PER_HOUR), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_add_months(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
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
        None,
        line,
    );
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_add_years(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let years_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, years_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, years_slot, line);
    push_const(chunk, Value::F64(12.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_datetime_add_months(chunks, current, line);
}

pub fn emit_datetime_days_in_month(chunks: &mut [Chunk], current: usize, line: u32) {
    // Shared arithmetic. Was `UTC(y, m, 1, …) - 86_400_000` then getUTCDate —
    // the day-0 rollover trick, correct but a host call plus a temporary Date
    // to answer a question about two integers. .NET months are 1-based.
    vybe_compiler::primitives::datetime::emit_days_in_month(
        &mut chunks[current],
        vybe_ast::datetime::MonthIndexing::OneBased,
        line,
    );
}
pub fn emit_datetime_is_leap_year(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let year_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, year_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, year_slot, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if(line);
    chunk.emit_struct_new(0, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const("Year must be between 1 and 9999.", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        chunk,
        "ArgumentOutOfRangeException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
    // Rule is shared; the range guard above is .NET's own (`DateTime.IsLeapYear`
    // throws for year < 1, which no other language does).
    chunk.emit_op_u16(Op::LOCAL_GET, year_slot, line);
    vybe_compiler::primitives::datetime::emit_is_leap_year(chunk, line);
    // The shared emitter yields an i32 0/1; .NET's surface is `bool`, which
    // renders "True"/"False" — lift it, exactly as Python's `isleap` does.
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
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

pub fn emit_datetime_add_timespan(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let span_slot = chunk.alloc_scratch(2);
    let date_slot = span_slot + 1;
    let total_ms_key = string_key(chunk, "TotalMilliseconds");
    chunk.emit_op_u16(Op::LOCAL_SET, span_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
    emit_datetime_time_from_obj(chunk, date_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, span_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, total_ms_key, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_subtract(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(3);
    let left_slot = right_slot + 1;
    let right_type_slot = right_slot + 2;
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    struct_get_named_field(chunk, TYPE_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_type_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_type_slot, line);
    push_const(chunk, Value::String(Arc::from("TimeSpan")), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    emit_datetime_time_from_obj(chunk, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    struct_get_named_field(chunk, "TotalMilliseconds", line);
    chunk.emit_op(Op::F64_SUB, line);
    emit_wrap_ms(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);
    emit_datetime_time_from_obj(chunk, left_slot, line);
    emit_datetime_time_from_obj(chunk, right_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    timespan_adapter::emit_build_timespan_from_total_ms(chunk, line);
    chunk.emit_end(line);
}

pub fn emit_datetime_equals(chunks: &mut [Chunk], current: usize, line: u32) {
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
    chunk.emit_op_u16(Op::LOCAL_GET, left_time_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_time_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
}

pub fn emit_datetime_to_binary(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_datetime_from_binary(chunks: &mut [Chunk], current: usize, line: u32) {
    let _ = (chunks, current, line);
}

pub fn emit_datetime_to_file_time_utc(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_datetime_time_from_obj(chunk, obj_slot, line);
}

pub fn emit_datetime_from_file_time_utc(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_wrap_ms(chunks, current, line);
    let chunk = &mut chunks[current];
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Utc")), line);
    struct_set_named_field_drop(chunk, "Kind", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Utc")), line);
    struct_set_named_field_drop(chunk, "kind", line);
}

pub fn emit_datetime_to_oadate(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_datetime_time_from_obj(chunk, obj_slot, line);
}

pub fn emit_datetime_from_oadate(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_from_binary(chunks, current, line);
}

pub fn emit_datetime_get_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_datetime_time_from_obj(chunk, obj_slot, line);
}

pub fn emit_datetime_to_universal_time(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_datetime_time_from_obj(chunk, obj_slot, line);
    emit_wrap_ms(chunks, current, line);
    let chunk = &mut chunks[current];
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Utc")), line);
    struct_set_named_field_drop(chunk, "Kind", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Utc")), line);
    struct_set_named_field_drop(chunk, "kind", line);
}

pub fn emit_datetime_to_local_time(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_datetime_time_from_obj(chunk, obj_slot, line);
    emit_wrap_ms(chunks, current, line);
    let chunk = &mut chunks[current];
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Local")), line);
    struct_set_named_field_drop(chunk, "Kind", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Local")), line);
    struct_set_named_field_drop(chunk, "kind", line);
}

pub fn emit_datetime_specify_kind(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let kind_slot = chunk.alloc_scratch(2);
    let date_slot = kind_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, kind_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
    emit_datetime_time_from_obj(chunk, date_slot, line);
    emit_wrap_ms(chunks, current, line);
    let chunk = &mut chunks[current];
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    struct_set_named_field_drop(chunk, "Kind", line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    struct_set_named_field_drop(chunk, "kind", line);
}

fn emit_iso_string_from_datetime_obj(chunks: &mut [Chunk], current: usize, obj_slot: u16, line: u32) {
    emit_datetime_time_from_obj(&mut chunks[current], obj_slot, line);
    call_import(chunks, current, "ecma:date", "toISOString", 1, line);
}

fn bind_datetime_to_string(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, line: u32) {
    let mut method = create_function_chunk("__datetime_tostring", 1);
    let time_key = method.add_constant(Value::String(Arc::from(TIME_KEY)));
    method.emit_op_u16(Op::LOCAL_GET, 0, line);
    method.emit_struct_field_op(Op::STRUCT_GET, 0, time_key, line);
    let iso = method.add_import("ecma:date", "toISOString");
    method.emit_call(iso, 1, line);
    method.emit_i32_const(0, line);
    method.emit_i32_const(19, line);
    let substring = method.add_import("wasm:js-string", "substring");
    method.emit_call(substring, 3, line);
    method.emit_op(Op::RETURN, line);
    method.local_count = 1;
    chunks.push(method);
    let method_idx = chunks.len() - 1;

    emit_bind_method(
        &mut chunks[current],
        obj_slot,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString),
        method_idx,
        line,
    );
}

fn bind_datetime_date_to_string(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, line: u32) {
    let mut method = create_function_chunk("__datetime_date_tostring", 1);
    let time_key = method.add_constant(Value::String(Arc::from(TIME_KEY)));
    method.emit_op_u16(Op::LOCAL_GET, 0, line);
    method.emit_struct_field_op(Op::STRUCT_GET, 0, time_key, line);
    let iso = method.add_import("ecma:date", "toISOString");
    method.emit_call(iso, 1, line);
    method.emit_i32_const(0, line);
    method.emit_i32_const(10, line);
    let substring = method.add_import("wasm:js-string", "substring");
    method.emit_call(substring, 3, line);
    method.emit_op(Op::RETURN, line);
    method.local_count = 1;
    chunks.push(method);
    let method_idx = chunks.len() - 1;

    emit_bind_method(
        &mut chunks[current],
        obj_slot,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString),
        method_idx,
        line,
    );
}

fn emit_substring_from_slot(chunk: &mut Chunk, str_slot: u16, start: i32, end: i32, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, str_slot, line);
    push_const(chunk, Value::I32(start), line);
    push_const(chunk, Value::I32(end), line);
    let substring = chunk.add_import("wasm:js-string", "substring");
    chunk.emit_call(substring, 3, line);
}

fn emit_concat(chunk: &mut Chunk, line: u32) {
    let concat = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(concat, 2, line);
}

pub fn emit_datetime_to_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(3);
    let format_slot = obj_slot + 1;
    let iso_slot = obj_slot + 2;
    if argc >= 3 {
        chunk.emit_op(Op::DROP, line);
    }
    if argc >= 2 {
        chunk.emit_op_u16(Op::LOCAL_SET, format_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_iso_string_from_datetime_obj(chunks, current, obj_slot, line);
    if argc >= 2 {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, iso_slot, line);

        chunk.emit_op_u16(Op::LOCAL_GET, format_slot, line);
        push_const(chunk, Value::String(Arc::from("yyyy/MM/dd")), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        emit_substring_from_slot(chunk, iso_slot, 0, 4, line);
        push_const(chunk, Value::String(Arc::from("/")), line);
        emit_concat(chunk, line);
        emit_substring_from_slot(chunk, iso_slot, 5, 7, line);
        emit_concat(chunk, line);
        push_const(chunk, Value::String(Arc::from("/")), line);
        emit_concat(chunk, line);
        emit_substring_from_slot(chunk, iso_slot, 8, 10, line);
        emit_concat(chunk, line);
        chunk.emit_else(line);

        chunk.emit_op_u16(Op::LOCAL_GET, format_slot, line);
        push_const(chunk, Value::String(Arc::from("yyyy-MM-dd")), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        emit_substring_from_slot(chunk, iso_slot, 0, 10, line);
        chunk.emit_else(line);

        chunk.emit_op_u16(Op::LOCAL_GET, format_slot, line);
        push_const(chunk, Value::String(Arc::from("yyyy-MM-dd HH:mm")), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        emit_substring_from_slot(chunk, iso_slot, 0, 10, line);
        push_const(chunk, Value::String(Arc::from(" ")), line);
        emit_concat(chunk, line);
        emit_substring_from_slot(chunk, iso_slot, 11, 16, line);
        emit_concat(chunk, line);
        chunk.emit_else(line);
        emit_substring_from_slot(chunk, iso_slot, 0, 19, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }
}

pub fn emit_datetime_parse_exact(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let input_slot = chunk.alloc_scratch(1);
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    call_import(chunks, current, "ecma:date", "parse", 1, line);
    emit_wrap_ms(chunks, current, line);
}

fn emit_timespan_total_ms_from_obj(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = string_key(chunk, "TotalMilliseconds");
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

fn bind_datetimeoffset_roles(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, line: u32) {
    let mut eq_method = create_function_chunk("__datetimeoffset_eq", 2);
    let time_key = eq_method.add_constant(Value::String(Arc::from(TIME_KEY)));
    eq_method.emit_op_u16(Op::LOCAL_GET, 1, line);
    eq_method.emit_op(Op::REF_IS_NULL, line);
    eq_method.emit_if(line);
    eq_method.emit_bool_const(false, line);
    eq_method.emit_else(line);
    eq_method.emit_op_u16(Op::LOCAL_GET, 0, line);
    eq_method.emit_struct_field_op(Op::STRUCT_GET, 0, time_key, line);
    eq_method.emit_op_u16(Op::LOCAL_GET, 1, line);
    eq_method.emit_struct_field_op(Op::STRUCT_GET, 0, time_key, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut eq_method, line);
    emit_i32_condition_as_bool(&mut eq_method, line);
    eq_method.emit_end(line);
    eq_method.emit_op(Op::RETURN, line);
    eq_method.local_count = 2;
    chunks.push(eq_method);
    let eq_idx = chunks.len() - 1;

    emit_bind_method(
        &mut chunks[current],
        obj_slot,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Eq),
        eq_idx,
        line,
    );

    let mut tostring_method = create_function_chunk("__datetimeoffset_tostring", 1);
    let datetime_key = tostring_method.add_constant(Value::String(Arc::from("DateTime")));
    let time_key = tostring_method.add_constant(Value::String(Arc::from(TIME_KEY)));
    tostring_method.emit_op_u16(Op::LOCAL_GET, 0, line);
    tostring_method.emit_struct_field_op(Op::STRUCT_GET, 0, datetime_key, line);
    tostring_method.emit_struct_field_op(Op::STRUCT_GET, 0, time_key, line);
    let iso = tostring_method.add_import("ecma:date", "toISOString");
    tostring_method.emit_call(iso, 1, line);
    tostring_method.emit_i32_const(0, line);
    tostring_method.emit_i32_const(19, line);
    let substring = tostring_method.add_import("wasm:js-string", "substring");
    tostring_method.emit_call(substring, 3, line);
    tostring_method.emit_op(Op::RETURN, line);
    tostring_method.local_count = 1;
    chunks.push(tostring_method);
    let tostring_idx = chunks.len() - 1;

    emit_bind_method(
        &mut chunks[current],
        obj_slot,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString),
        tostring_idx,
        line,
    );
}

fn emit_datetimeoffset_wrap_utc_offset(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let offset_ms_slot = chunk.alloc_scratch(4);
    let utc_ms_slot = offset_ms_slot + 1;
    let local_dt_slot = offset_ms_slot + 2;
    let obj_slot = offset_ms_slot + 3;
    chunk.emit_op_u16(Op::LOCAL_SET, offset_ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, utc_ms_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, utc_ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, offset_ms_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_wrap_ms(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, local_dt_slot, line);

    call_import(chunks, current, "ecma:object", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("DateTimeOffset")), line);
    struct_set_named_field_drop(chunk, TYPE_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, utc_ms_slot, line);
    struct_set_named_field_drop(chunk, TIME_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, offset_ms_slot, line);
    struct_set_named_field_drop(chunk, "__offset_ms", line);

    for field in [
        "Year",
        "Month",
        "Day",
        "Hour",
        "Minute",
        "Second",
        "Millisecond",
        "DayOfYear",
        "Ticks",
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        emit_named_field_from_obj(chunk, local_dt_slot, field, line);
        struct_set_named_field_drop(chunk, field, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, local_dt_slot, line);
    struct_set_named_field_drop(chunk, "DateTime", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    emit_named_field_from_obj(chunk, local_dt_slot, "Date", line);
    let date_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
    bind_datetime_date_to_string(chunks, current, date_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, date_slot, line);
    struct_set_named_field_drop(chunk, "Date", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, offset_ms_slot, line);
    timespan_adapter::emit_build_timespan_from_total_ms(chunk, line);
    struct_set_named_field_drop(chunk, "Offset", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, utc_ms_slot, line);
    emit_wrap_ms(chunks, current, line);
    let chunk = &mut chunks[current];
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Utc")), line);
    struct_set_named_field_drop(chunk, "Kind", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Utc")), line);
    struct_set_named_field_drop(chunk, "kind", line);
    struct_set_named_field_drop(chunk, "UtcDateTime", line);

    bind_datetimeoffset_roles(chunks, current, obj_slot, line);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_datetimeoffset_new(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let offset_obj_slot = chunk.alloc_scratch(8);
    let second_slot = offset_obj_slot + 1;
    let minute_slot = offset_obj_slot + 2;
    let hour_slot = offset_obj_slot + 3;
    let day_slot = offset_obj_slot + 4;
    let month_slot = offset_obj_slot + 5;
    let year_slot = offset_obj_slot + 6;
    let offset_ms_slot = offset_obj_slot + 7;
    chunk.emit_op_u16(Op::LOCAL_SET, offset_obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, second_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, minute_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hour_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, day_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, month_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, year_slot, line);
    emit_timespan_total_ms_from_obj(chunk, offset_obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, offset_ms_slot, line);
    emit_utc_from_slots(
        chunks,
        current,
        year_slot,
        month_slot,
        day_slot,
        Some(hour_slot),
        Some(minute_slot),
        Some(second_slot),
        None,
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, offset_ms_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_GET, offset_ms_slot, line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
}

pub fn emit_datetimeoffset_utc_now(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    call_import(chunks, current, "ecma:date", "now", 0, line);
    push_const(&mut chunks[current], Value::F64(0.0), line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
}

pub fn emit_datetimeoffset_min_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    push_const(
        &mut chunks[current],
        Value::F64(vybe_compiler::primitives::datetime::DOTNET_DATETIME_MIN_UNIX_MS),
        line,
    );
    push_const(&mut chunks[current], Value::F64(0.0), line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
}

pub fn emit_datetimeoffset_max_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    push_const(
        &mut chunks[current],
        Value::F64(vybe_compiler::primitives::datetime::DOTNET_DATETIME_MAX_UNIX_MS),
        line,
    );
    push_const(&mut chunks[current], Value::F64(0.0), line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
}

pub fn emit_datetimeoffset_from_unix_time_seconds(
    chunks: &mut Vec<Chunk>,
    current: usize,
    line: u32,
) {
    push_const(
        &mut chunks[current],
        Value::F64(vybe_compiler::primitives::datetime::MS_PER_SECOND),
        line,
    );
    chunks[current].emit_op(Op::F64_MUL, line);
    push_const(&mut chunks[current], Value::F64(0.0), line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
}

pub fn emit_datetimeoffset_from_unix_time_milliseconds(
    chunks: &mut Vec<Chunk>,
    current: usize,
    line: u32,
) {
    push_const(&mut chunks[current], Value::F64(0.0), line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
}

pub fn emit_datetimeoffset_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    call_import(chunks, current, "ecma:date", "parse", 1, line);
    push_const(&mut chunks[current], Value::F64(0.0), line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
}

pub fn emit_datetimeoffset_try_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    call_import(chunks, current, "ecma:date", "parse", 1, line);
    let chunk = &mut chunks[current];
    let ms_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_datetimeoffset_add_hours(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let hours_slot = chunk.alloc_scratch(2);
    let obj_slot = hours_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, hours_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_named_field_from_obj(chunk, obj_slot, TIME_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_GET, hours_slot, line);
    push_const(chunk, Value::F64(vybe_compiler::primitives::datetime::MS_PER_HOUR), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_named_field_from_obj(chunk, obj_slot, "__offset_ms", line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
}

pub fn emit_datetimeoffset_to_universal_time(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], TIME_KEY, line);
    push_const(&mut chunks[current], Value::F64(0.0), line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
}

pub fn emit_datetimeoffset_to_offset(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let offset_obj_slot = chunk.alloc_scratch(2);
    let obj_slot = offset_obj_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, offset_obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_named_field_from_obj(chunk, obj_slot, TIME_KEY, line);
    emit_timespan_total_ms_from_obj(chunk, offset_obj_slot, line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
}

pub fn emit_datetimeoffset_to_unix_time_seconds(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], TIME_KEY, line);
    push_const(
        &mut chunks[current],
        Value::F64(vybe_compiler::primitives::datetime::MS_PER_SECOND),
        line,
    );
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
}

pub fn emit_datetimeoffset_to_unix_time_milliseconds(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
) {
    struct_get_named_field(&mut chunks[current], TIME_KEY, line);
}

pub fn emit_datetimeoffset_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(2);
    let left_slot = right_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    emit_named_field_from_obj(chunk, left_slot, TIME_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    emit_named_field_from_obj(chunk, right_slot, TIME_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    emit_compare_numeric_slots(chunk, left_slot, right_slot, line);
}

pub fn emit_datetimeoffset_subtract(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(2);
    let left_slot = right_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    emit_named_field_from_obj(chunk, left_slot, TIME_KEY, line);
    emit_named_field_from_obj(chunk, right_slot, TIME_KEY, line);
    chunk.emit_op(Op::F64_SUB, line);
    timespan_adapter::emit_build_timespan_from_total_ms(chunk, line);
}

pub fn emit_datetimeoffset_equals(chunks: &mut [Chunk], current: usize, exact: bool, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = chunk.alloc_scratch(2);
    let left_slot = right_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    emit_named_field_from_obj(chunk, left_slot, TIME_KEY, line);
    emit_named_field_from_obj(chunk, right_slot, TIME_KEY, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    if exact {
        emit_named_field_from_obj(chunk, left_slot, "__offset_ms", line);
        emit_named_field_from_obj(chunk, right_slot, "__offset_ms", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_op(Op::I32_AND, line);
    }
    emit_i32_condition_as_bool(chunk, line);
}

pub fn emit_datetimeoffset_get_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], TIME_KEY, line);
}

fn emit_offset_text_from_slot(chunk: &mut Chunk, offset_ms_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, offset_ms_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::String(Arc::from("+")), line);
    chunk.emit_else(line);
    push_const(chunk, Value::String(Arc::from("-")), line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, offset_ms_slot, line);
    chunk.emit_op(Op::F64_ABS, line);
    push_const(
        chunk,
        Value::F64(vybe_compiler::primitives::datetime::MS_PER_HOUR),
        line,
    );
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_TRUNC, line);
    push_const(chunk, Value::I32(10), line);
    let number_to_string = chunk.add_import("ecma:number", "toString");
    chunk.emit_call(number_to_string, 2, line);
    push_const(chunk, Value::I32(2), line);
    push_const(chunk, Value::String(Arc::from("0")), line);
    let pad_start = chunk.add_import("ecma:string", "padStart");
    chunk.emit_call(pad_start, 3, line);
    emit_concat(chunk, line);
    push_const(chunk, Value::String(Arc::from(":00")), line);
    emit_concat(chunk, line);
}

pub fn emit_datetimeoffset_to_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let obj_slot = chunks[current].alloc_scratch(5);
    let format_slot = obj_slot + 1;
    let iso_slot = obj_slot + 2;
    let offset_ms_slot = obj_slot + 3;
    let local_dt_slot = obj_slot + 4;

    {
        let chunk = &mut chunks[current];
        if argc >= 3 {
            chunk.emit_op(Op::DROP, line);
        }
        if argc >= 2 {
            chunk.emit_op_u16(Op::LOCAL_SET, format_slot, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

        emit_named_field_from_obj(chunk, obj_slot, "DateTime", line);
        chunk.emit_op_u16(Op::LOCAL_SET, local_dt_slot, line);
    }
    emit_iso_string_from_datetime_obj(chunks, current, local_dt_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, iso_slot, line);
    emit_named_field_from_obj(chunk, obj_slot, "__offset_ms", line);
    chunk.emit_op_u16(Op::LOCAL_SET, offset_ms_slot, line);

    if argc >= 2 {
        chunk.emit_op_u16(Op::LOCAL_GET, format_slot, line);
        push_const(chunk, Value::String(Arc::from("yyyy-MM-dd HH:mm zzz")), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        emit_substring_from_slot(chunk, iso_slot, 0, 10, line);
        push_const(chunk, Value::String(Arc::from(" ")), line);
        emit_concat(chunk, line);
        emit_substring_from_slot(chunk, iso_slot, 11, 16, line);
        emit_concat(chunk, line);
        push_const(chunk, Value::String(Arc::from(" ")), line);
        emit_concat(chunk, line);
        emit_offset_text_from_slot(chunk, offset_ms_slot, line);
        emit_concat(chunk, line);
        chunk.emit_else(line);

        chunk.emit_op_u16(Op::LOCAL_GET, format_slot, line);
        push_const(chunk, Value::String(Arc::from("o")), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        emit_substring_from_slot(chunk, iso_slot, 0, 19, line);
        emit_offset_text_from_slot(chunk, offset_ms_slot, line);
        emit_concat(chunk, line);
        chunk.emit_else(line);
        emit_substring_from_slot(chunk, iso_slot, 0, 19, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    } else {
        emit_substring_from_slot(chunk, iso_slot, 0, 19, line);
    }
}
