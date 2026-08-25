//! .NET `System.DateTime` adapter — bytecode-only.
//!
//! `DateTime` is a calendar instant; .NET's `DateTime.Now /
//! UtcNow / Parse / Today` static methods plus the `New
//! DateTime(y, m, d, ...)` constructor produce a value with
//! `__type=DateTime` and a millisecond-since-epoch payload.
//!
//! The underlying primitive is `wasi:clocks/system-clock.now`
//! (`wasi:clocks@0.3.1`; 0.2 spelled the interface `wall-clock`).
//! `ecma:date.now` reads through it and
//! returns ms since epoch — the form ECMA-262 §21.4 [[DateValue]]
//! uses. Each adapter wraps that ms in a DateTime-shaped Object so
//! the .NET surface looks .NET-shaped while the bytecode is
//! standardized.

use std::sync::Arc;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::object::{emit_bind_method, emit_bind_method_with_slot};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

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
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
}

// struct.set is spec-shaped now (pushes nothing) — the old `_drop` variant
// compensated for the retired push-val-back contract.
fn struct_set_named_field_drop(chunk: &mut Chunk, key: &str, line: u32) {
    struct_set_named_field(chunk, key, line);
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

/// The millisecond payload of the DateTime object held in `obj_slot`, and the
/// named field beside it — the two accessors sibling adapters need so that
/// `__time` is spelled in exactly one place.
///
/// Stack: `[]` → `[ms]` / `[]` → `[field]`.
pub fn emit_millis_from_slot(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    emit_datetime_time_from_obj(chunk, obj_slot, line);
}

pub fn emit_field_from_slot(chunk: &mut Chunk, obj_slot: u16, field: &str, line: u32) {
    emit_named_field_from_obj(chunk, obj_slot, field, line);
}

/// `-1` / `0` / `1` from two numeric slots.
///
/// ⛔ Both `if`s were `emit_if`, i.e. `emit_if_params(line, 0, 0)` — VOID — so
/// the `-1` and `1` pushed inside them were DISCARDED and
/// `DateTime.Compare(a, b)` answered `0` for every pair. Identical to the
/// `Version.CompareTo` defect; both stayed hidden because the VB walker folded
/// literal comparisons at compile time.
fn emit_compare_numeric_slots(chunk: &mut Chunk, left_slot: u16, right_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::I32(-1), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if_value(line);
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
    // The enum's NUMERIC value. `Value`/`value` hold the NAME, because that is
    // what `CStr(d.DayOfWeek)` must print, so a consumer that needs the index —
    // `datetime_format_adapter`, for `ddd`/`dddd` — had nothing to read and
    // multiplied a string by the name width.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, dow_slot, line);
    struct_set_named_field_drop(chunk, "__index", line);
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
        chunks, current, year_slot, one_slot, one_slot, None, None, None, None, line,
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

    // ⛔ BOTH spellings, like `Kind`/`kind` above and the component loop below:
    // a case-insensitive frontend folds the member name, so a PascalCase-only
    // field is invisible to it — and reads `undefined` rather than erroring.
    for spelling in ["Ticks", "ticks"] {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
        push_const(chunk, Value::F64(10_000.0), line);
        chunk.emit_op(Op::F64_MUL, line);
        push_const(chunk, Value::F64(621_355_968_000_000_000.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        struct_set_named_field_drop(chunk, spelling, line);
    }

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

    // Both spellings — `DayOfYear` was PascalCase-only, the same silent gap
    // `Ticks` had.
    for spelling in ["DayOfYear", "dayofyear"] {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, year_start_slot, line);
        chunk.emit_op(Op::F64_SUB, line);
        push_const(
            chunk,
            Value::F64(vybe_compiler::primitives::datetime::MS_PER_DAY),
            line,
        );
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_TRUNC, line);
        push_const(chunk, Value::I32(1), line);
        chunk.emit_op(Op::F64_ADD, line);
        struct_set_named_field_drop(chunk, spelling, line);
    }

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
        push_const(
            chunk,
            Value::F64(vybe_compiler::primitives::datetime::MS_PER_HOUR),
            line,
        );
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
        push_const(
            chunk,
            Value::F64(vybe_compiler::primitives::datetime::MS_PER_HOUR),
            line,
        );
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
    bind_datetime_compare(chunks, current, obj_slot, line);
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
/// reads through `wasi:clocks/system-clock.now`) and wrap in a
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
                if argc >= 7 {
                    Some(millis_or_kind_slot)
                } else {
                    None
                },
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
    push_const(
        chunk,
        Value::F64(vybe_compiler::primitives::datetime::MS_PER_DAY),
        line,
    );
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
    push_const(
        chunk,
        Value::F64(vybe_compiler::primitives::datetime::MS_PER_HOUR),
        line,
    );
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_wrap_ms(chunks, current, line);
}

pub fn emit_datetime_add_months(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let months_slot = chunk.alloc_scratch(10);
    let date_slot = months_slot + 1;
    let year_slot = months_slot + 2;
    let month_slot = months_slot + 3;
    let day_slot = months_slot + 4;
    let hour_slot = months_slot + 5;
    let minute_slot = months_slot + 6;
    let second_slot = months_slot + 7;
    let total_months_slot = months_slot + 8;
    let dim_slot = months_slot + 9;
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

    // .NET CLAMPS (`MonthOverflow::Clamp` — MEASURED against `dotnet run`:
    // 2024-02-29 AddYears(1) → 2025-02-28). The day passed through to
    // `Date.UTC` unchanged, whose normalising is OVERFLOW semantics — the
    // PHP/JS policy, not .NET's — so Feb 29 rolled into Mar 1. Clamp to the
    // target month's length first; the rule is the shared primitive.
    chunk.emit_op_u16(Op::LOCAL_GET, year_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, month_slot, line);
    vybe_compiler::primitives::datetime::emit_days_in_month(
        chunk,
        vybe_ast::datetime::MonthIndexing::OneBased,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, dim_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, day_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, dim_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, dim_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, day_slot, line);
    chunk.emit_end(line);

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

/// `a.CompareTo(b)` — negative when `a` precedes `b`. **Instance shape**: the
/// receiver arrives BENEATH the argument, so the first `local.set` takes `b`.
///
/// ⛔ The static `DateTime.Compare(a, b)` declared on the same component class
/// points at this emitter but delivers its operands in the OPPOSITE order —
/// measured both ways: with the order below, `a.CompareTo(b)` and every
/// relational operator derived from it are correct while
/// `Date.Compare(#2024-01-01#, #2024-01-02#)` answers `+1`; exchanging the two
/// `local.set`s inverts exactly that set (5 relational tests break, the static
/// one passes). Hence [`emit_datetime_compare_static`], rather than one order
/// that is wrong for half its callers. Nothing could see either bug while
/// `emit_compare_numeric_slots` discarded its result into a void `if` and
/// returned `0` for every pair.
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

/// Static `DateTime.Compare(a, b)` — the same comparison as
/// [`emit_datetime_compare`].
///
/// ⛔ Do NOT exchange the operands. Static and instance shapes BOTH push
/// left-to-right, and `emit_datetime_compare` pops `right` then `left`, which
/// is correct for each; a swap here inverts every static answer.
pub fn emit_datetime_compare_static(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_datetime_compare(chunks, current, line);
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

fn emit_iso_string_from_datetime_obj(
    chunks: &mut [Chunk],
    current: usize,
    obj_slot: u16,
    line: u32,
) {
    emit_datetime_time_from_obj(&mut chunks[current], obj_slot, line);
    call_import(chunks, current, "ecma:date", "toISOString", 1, line);
}

/// Stack: `[]` → `[str]` — the field named `field` of the object in
/// `obj_slot`, as a string.
fn emit_field_as_string(chunk: &mut Chunk, obj_slot: u16, field: &str, line: u32) {
    emit_named_field_from_obj(chunk, obj_slot, field, line);
    let to_str = chunk.add_import("ecma:string", "String");
    chunk.emit_call(to_str, 1, line);
}

/// Stack: `[]` → `[str]` — `field` as a string, zero-padded to two digits.
fn emit_field_padded(chunk: &mut Chunk, obj_slot: u16, field: &str, line: u32) {
    emit_named_field_from_obj(chunk, obj_slot, field, line);
    push_const(chunk, Value::F64(10.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::String(Arc::from("0")), line);
    emit_field_as_string(chunk, obj_slot, field, line);
    emit_concat(chunk, line);
    chunk.emit_else(line);
    emit_field_as_string(chunk, obj_slot, field, line);
    chunk.emit_end(line);
}

/// `DateTime.ToString()` — the .NET general format `"M/d/yyyy h:mm:ss tt"`.
///
/// This used to be `toISOString()[0..19]`, i.e. `2024-05-14T15:45:59`. Real
/// .NET renders `5/14/2024 3:45:59 PM`, and so does `dotnet run`. The
/// deviation stayed invisible because VB's walker folded every date to its
/// display text at compile time and never reached this method — one more leaf
/// defect that a language-local copy was standing in front of.
///
/// `elide_midnight` is VB's `CStr` rule: `CStr` on a Date whose time is
/// exactly midnight prints the date alone (`5/19/2024`). C#'s
/// `DateTime.ToString()` always prints the time (`5/19/2024 12:00:00 AM`).
/// Both spellings currently share ONE `ToString` slot, so the binding below
/// passes `true` — VB's rule — and splitting them needs `CStr` to bind its own
/// method rather than reading the object's `ToString`. Left as one knob here
/// so the difference is visible at the call site instead of being a surprise.
///
/// Stack: `[]` → `[str]`.
fn emit_datetime_display(chunk: &mut Chunk, obj_slot: u16, elide_midnight: bool, line: u32) {
    if elide_midnight {
        // VB's `CStr` prints only the half that carries information: a Date at
        // midnight is `5/19/2024`, and a value on `DateTime.MinValue`'s date —
        // which is where `TimeSerial` puts a time of day — is `9:30:00 AM`.
        emit_named_field_from_obj(chunk, obj_slot, "Year", line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_EQ, line);
        chunk.emit_if_value(line);
        emit_time_of_day_suffix(chunk, obj_slot, false, line);
        chunk.emit_else(line);
        emit_date_half(chunk, obj_slot, line);
        emit_seconds_of_day(chunk, obj_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        chunk.emit_op(Op::F64_NE, line);
        chunk.emit_if_value(line);
        emit_time_of_day_suffix(chunk, obj_slot, true, line);
        chunk.emit_else(line);
        push_const(chunk, Value::String(Arc::from("")), line);
        chunk.emit_end(line);
        emit_concat(chunk, line);
        chunk.emit_end(line);
    } else {
        emit_date_half(chunk, obj_slot, line);
        emit_time_of_day_suffix(chunk, obj_slot, true, line);
        emit_concat(chunk, line);
    }
}

/// `"M/d/yyyy"`. Stack: `[]` → `[str]`.
fn emit_date_half(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    emit_field_as_string(chunk, obj_slot, "Month", line);
    push_const(chunk, Value::String(Arc::from("/")), line);
    emit_concat(chunk, line);
    emit_field_as_string(chunk, obj_slot, "Day", line);
    emit_concat(chunk, line);
    push_const(chunk, Value::String(Arc::from("/")), line);
    emit_concat(chunk, line);
    emit_field_as_string(chunk, obj_slot, "Year", line);
    emit_concat(chunk, line);
}

/// Seconds since midnight. Stack: `[]` → `[n]`.
fn emit_seconds_of_day(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    emit_named_field_from_obj(chunk, obj_slot, "Hour", line);
    push_const(chunk, Value::F64(60.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_named_field_from_obj(chunk, obj_slot, "Minute", line);
    chunk.emit_op(Op::F64_ADD, line);
    push_const(chunk, Value::F64(60.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_named_field_from_obj(chunk, obj_slot, "Second", line);
    chunk.emit_op(Op::F64_ADD, line);
}

/// `"h:mm:ss tt"`, with a leading space when it follows a date.
///
/// Stack: `[]` → `[str]`.
fn emit_time_of_day_suffix(chunk: &mut Chunk, obj_slot: u16, leading_space: bool, line: u32) {
    push_const(
        chunk,
        Value::String(Arc::from(if leading_space { " " } else { "" })),
        line,
    );

    // 12-hour clock: `Hour % 12`, with 0 displayed as 12.
    emit_named_field_from_obj(chunk, obj_slot, "Hour", line);
    push_const(chunk, Value::F64(12.0), line);
    vybe_compiler::primitives::math::emit_c_fmod(chunk, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::String(Arc::from("12")), line);
    chunk.emit_else(line);
    emit_named_field_from_obj(chunk, obj_slot, "Hour", line);
    push_const(chunk, Value::F64(12.0), line);
    vybe_compiler::primitives::math::emit_c_fmod(chunk, line);
    let to_str = chunk.add_import("ecma:string", "String");
    chunk.emit_call(to_str, 1, line);
    chunk.emit_end(line);
    emit_concat(chunk, line);

    push_const(chunk, Value::String(Arc::from(":")), line);
    emit_concat(chunk, line);
    emit_field_padded(chunk, obj_slot, "Minute", line);
    emit_concat(chunk, line);
    push_const(chunk, Value::String(Arc::from(":")), line);
    emit_concat(chunk, line);
    emit_field_padded(chunk, obj_slot, "Second", line);
    emit_concat(chunk, line);

    push_const(chunk, Value::String(Arc::from(" ")), line);
    emit_concat(chunk, line);
    emit_named_field_from_obj(chunk, obj_slot, "Hour", line);
    push_const(chunk, Value::F64(12.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::String(Arc::from("AM")), line);
    chunk.emit_else(line);
    push_const(chunk, Value::String(Arc::from("PM")), line);
    chunk.emit_end(line);
    emit_concat(chunk, line);
}

/// Build the shared `ToString` body for a DateTime object. `elide_midnight`
/// selects VB's `CStr` rule — see [`emit_datetime_display`].
pub fn push_datetime_display_chunk(
    chunks: &mut Vec<Chunk>,
    name: &str,
    elide_midnight: bool,
    line: u32,
) -> usize {
    let mut method = create_function_chunk(name, 1);
    method.local_count = 1;
    emit_datetime_display(&mut method, 0, elide_midnight, line);
    method.emit_op(Op::RETURN, line);
    chunks.push(method);
    chunks.len() - 1
}

/// Bind `CompareTo` under [`ProtocolSlot::Compare`], which is how `<`, `<=`,
/// `>`, `>=` and `=` reach a DateTime at all.
///
/// The object bound only `ToString`, so every relational operator on two dates
/// fell through to numeric coercion and trapped with
/// `wasm:js-number.toF64 — not a number`. It stayed invisible while the VB
/// walker folded `#…# < #…#` to a boolean at compile time — the same shape as
/// the `Version` compare defect, on the adjacent type.
///
/// `emit_rich_compare_locals` derives every relational operator from the sign
/// of this one method, so one binding answers all six.
fn bind_datetime_compare(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, line: u32) {
    let mut method = create_function_chunk("__datetime_compareto", 2);
    method.local_count = 2;
    let time_key = method.add_constant(Value::String(Arc::from(TIME_KEY)));

    // ⛔ A NULL operand is `Nothing`, and `Nothing` for a VALUE TYPE is its
    // DEFAULT — `DateTime.MinValue` — not "no answer". Reading `__time` off
    // null yielded `undefined`, both `<` and `>` answered false, and the
    // method returned 0: "equal". Every `d = Nothing` was therefore TRUE and
    // every `d <> Nothing` FALSE, which is what made `DateTime.TryParse(s, d)`
    // report failure on a date it had just parsed correctly —
    // `lowering::try_parse_desugar` asks exactly that question.
    //
    // ⛔ It hid behind a walk-time fold: on a LITERAL the VB walker computes
    // the comparison itself and answers correctly, so the defect only shows
    // with a value the walker cannot see through.
    let load_time = move |method: &mut Chunk, local: u16| {
        method.emit_op_u16(Op::LOCAL_GET, local, line);
        method.emit_op(Op::REF_IS_NULL, line);
        method.emit_if_value(line);
        push_const(
            method,
            Value::F64(vybe_compiler::primitives::datetime::DOTNET_DATETIME_MIN_UNIX_MS),
            line,
        );
        method.emit_else(line);
        method.emit_op_u16(Op::LOCAL_GET, local, line);
        method.emit_struct_field_op(Op::STRUCT_GET, 0, time_key, line);
        method.emit_end(line);
    };

    load_time(&mut method, 0);
    load_time(&mut method, 1);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut method, line);
    method.emit_if_value(line);
    push_const(&mut method, Value::I32(-1), line);
    method.emit_else(line);
    load_time(&mut method, 0);
    load_time(&mut method, 1);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut method, line);
    method.emit_if_value(line);
    push_const(&mut method, Value::I32(1), line);
    method.emit_else(line);
    push_const(&mut method, Value::I32(0), line);
    method.emit_end(line);
    method.emit_end(line);
    method.emit_op(Op::RETURN, line);

    chunks.push(method);
    let method_idx = chunks.len() - 1;
    for name in ["CompareTo", "compareto", "compare"] {
        emit_bind_method_with_slot(
            &mut chunks[current],
            obj_slot,
            name,
            Some(vybe_ast::ProtocolSlot::Compare),
            method_idx,
            None,
            line,
        );
    }
}

fn bind_datetime_to_string(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, line: u32) {
    let method_idx =
        push_datetime_display_chunk(chunks, "__datetime_tostring", true, line);

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

/// `value.ToString([format[, provider]])`.
///
/// The pattern is interpreted by `datetime_format_adapter`. No format at all
/// is `"G"`, .NET's general pattern.
pub fn emit_datetime_to_string(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        let obj_slot = chunk.alloc_scratch(2);
        let fmt_slot = obj_slot + 1;
        // A provider argument beyond the format sits on TOP; the invariant
        // culture is the only one this surface renders.
        for _ in 2..argc {
            chunk.emit_op(Op::DROP, line);
        }
        if argc >= 2 {
            chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
        } else {
            push_const(chunk, Value::String(Arc::from("G")), line);
            chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    }
    super::datetime_format_adapter::emit_date_format(chunks, current, line);
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

    // `CompareTo` and the `Compare` SLOT the relational operators read. The
    // DateTimeOffset object stores `__time` as the UTC instant, which is
    // exactly what .NET compares two offsets by, so DateTime's body is the
    // right one — it was simply never bound here, leaving `a > b` and
    // `a.CompareTo(b)` unanswered on an offset.
    bind_datetime_compare(chunks, current, obj_slot, line);
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

    // ⛔ EVERY property in BOTH spellings, as `emit_wrap_ms_internal` writes
    // DateTime's. A case-insensitive frontend reads a member FOLDED — VB emits
    // `struct.get (hour)` for `dto.Hour` — so PascalCase alone is invisible.
    for field in [
        "Year",
        "Month",
        "Day",
        "Hour",
        "Minute",
        "Second",
        "Millisecond",
        "DayOfYear",
        "DayOfWeek",
        "TimeOfDay",
        "Ticks",
    ] {
        for spelling in [field.to_string(), field.to_ascii_lowercase()] {
            chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
            emit_named_field_from_obj(chunk, local_dt_slot, field, line);
            struct_set_named_field_drop(chunk, &spelling, line);
        }
    }

    // ⛔ NOT `LocalDateTime`. `DateTimeOffset.DateTime` is the offset-local
    // clock reading, which is what `local_dt` holds; `LocalDateTime` is
    // `UtcDateTime.ToLocalTime()` — the SYSTEM zone, Kind=Local — and the two
    // agree only when the system zone happens to equal the offset. Left
    // unregistered rather than aliased to the wrong value.
    for spelling in ["DateTime", "datetime"] {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, local_dt_slot, line);
        struct_set_named_field_drop(chunk, spelling, line);
    }
    emit_named_field_from_obj(chunk, local_dt_slot, "Date", line);
    let date_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
    bind_datetime_date_to_string(chunks, current, date_slot, line);
    let chunk = &mut chunks[current];
    for spelling in ["Date", "date"] {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, date_slot, line);
        struct_set_named_field_drop(chunk, spelling, line);
    }

    // ⛔ Each composite is BUILT ONCE into a slot, then written under both
    // spellings. Building per spelling runs `emit_wrap_ms` twice, and every
    // extra `alloc_scratch` here aliases the CALLER's named locals.
    let offset_ts_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, offset_ms_slot, line);
    timespan_adapter::emit_build_timespan_from_total_ms(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, offset_ts_slot, line);
    for spelling in ["Offset", "offset"] {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, offset_ts_slot, line);
        struct_set_named_field_drop(chunk, spelling, line);
    }

    let utc_dt_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, utc_ms_slot, line);
    emit_wrap_ms(chunks, current, line);
    let chunk = &mut chunks[current];
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Utc")), line);
    struct_set_named_field_drop(chunk, "Kind", line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Utc")), line);
    struct_set_named_field_drop(chunk, "kind", line);
    chunk.emit_op_u16(Op::LOCAL_SET, utc_dt_slot, line);
    for spelling in ["UtcDateTime", "utcdatetime"] {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, utc_dt_slot, line);
        struct_set_named_field_drop(chunk, spelling, line);
    }

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

/// `dto.Add<Unit>(n)` — shift the UTC instant by `n * unit_ms` and rebuild at
/// the SAME offset, which is what .NET does: the offset is a property of the
/// value, not of the instant.
///
/// The unit is the only thing that varies, so all six share this body.
fn emit_datetimeoffset_add_scaled(
    chunks: &mut Vec<Chunk>,
    current: usize,
    unit_ms: f64,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let amount_slot = chunk.alloc_scratch(2);
    let obj_slot = amount_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, amount_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_named_field_from_obj(chunk, obj_slot, TIME_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_GET, amount_slot, line);
    push_const(chunk, Value::F64(unit_ms), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    emit_named_field_from_obj(chunk, obj_slot, "__offset_ms", line);
    emit_datetimeoffset_wrap_utc_offset(chunks, current, line);
}

pub fn emit_datetimeoffset_add_hours(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_datetimeoffset_add_scaled(
        chunks,
        current,
        vybe_compiler::primitives::datetime::MS_PER_HOUR,
        line,
    );
}

pub fn emit_datetimeoffset_add_days(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_datetimeoffset_add_scaled(
        chunks,
        current,
        vybe_compiler::primitives::datetime::MS_PER_DAY,
        line,
    );
}

pub fn emit_datetimeoffset_add_minutes(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_datetimeoffset_add_scaled(chunks, current, 60_000.0, line);
}

pub fn emit_datetimeoffset_add_seconds(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_datetimeoffset_add_scaled(chunks, current, 1_000.0, line);
}

pub fn emit_datetimeoffset_add_milliseconds(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_datetimeoffset_add_scaled(chunks, current, 1.0, line);
}

/// `dto.AddTicks(n)` — 100-nanosecond units, .NET's own resolution.
pub fn emit_datetimeoffset_add_ticks(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_datetimeoffset_add_scaled(chunks, current, 0.000_1, line);
}

/// `dto.Add(timespan)`.
pub fn emit_datetimeoffset_add_timespan(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let ts_slot = chunk.alloc_scratch(2);
    let obj_slot = ts_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, ts_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_named_field_from_obj(chunk, obj_slot, TIME_KEY, line);
    emit_timespan_total_ms_from_obj(chunk, ts_slot, line);
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

/// `dto.ToString([format])` — the same interpreter as `DateTime`'s.
///
/// The offset object carries `Year`/`Month`/… and `__offset_ms`, which is all
/// `zzz` needs, so there is no second implementation to keep in step. Its
/// no-format rendering is .NET's general pattern plus the offset.
pub fn emit_datetimeoffset_to_string(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    {
        let chunk = &mut chunks[current];
        let obj_slot = chunk.alloc_scratch(2);
        let fmt_slot = obj_slot + 1;
        for _ in 2..argc {
            chunk.emit_op(Op::DROP, line);
        }
        if argc >= 2 {
            chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
        } else {
            push_const(
                chunk,
                Value::String(Arc::from("M/d/yyyy h:mm:ss tt zzz")),
                line,
            );
            chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    }
    super::datetime_format_adapter::emit_date_format(chunks, current, line);
}
