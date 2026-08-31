//! .NET `System.DateOnly` / `System.TimeOnly` adapters — bytecode-only.
//!
//! Both types were absent from the catalog entirely, so `new DateOnly(...)`
//! and `new TimeOnly(...)` reached no `ClassType` and trapped with
//! `undefined is not callable`.
//!
//! Neither type needs a new representation. `datetime_adapter` already turns a
//! millisecond timestamp into an object carrying every calendar component as a
//! named field in both spellings (`Year`/`year`, `Day`/`day`, …), and
//! `datetime_format_adapter` already renders `"d"` as `M/d/yyyy` and `"t"` as
//! `h:mm tt` — which is exactly what .NET prints for `DateOnly.ToString()` and
//! `TimeOnly.ToString()`. So both types ARE that object, minted through
//! [`datetime_adapter::emit_datetime_from_millis`] and then specialized:
//!
//! * `DateOnly` — the millisecond value floored to UTC midnight, plus
//!   `DayNumber` (days since `0001-01-01`), `ToString` re-bound to `"d"`.
//! * `TimeOnly` — the millisecond value reduced into `[0, 86_400_000)` ON the
//!   epoch day, so `getUTCHours` and friends already answer `Hour`/`Minute`/
//!   `Second`/`Millisecond`. `Ticks` is overwritten because .NET counts a
//!   TimeOnly's ticks FROM MIDNIGHT, not from `0001-01-01` like a DateTime's.
//!
//! ⛔ `TimeOnly.AddHours` / `AddMinutes` / `Add` are MODULAR — .NET wraps them
//! into the same day (`23:00 + 2h` is `01:00`, and `00:30 - 2h` is `22:30`).
//! They must NOT forward to the `DateTime` adds, which roll into the next day.
//! The corpus cannot see the difference — every case is `h` in 1..20 plus two
//! hours — so a gate that forwarded them would have been green and wrong.
//! Measured against `dotnet` 10 (`/usr/local/share/dotnet/dotnet`).

use std::sync::Arc;

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::datetime::{
    DOTNET_DATETIME_MAX_UNIX_MS, DOTNET_DATETIME_MIN_UNIX_MS, DOTNET_TICKS_PER_MS, MS_PER_DAY,
    MS_PER_HOUR, MS_PER_MINUTE, MS_PER_SECOND,
};
use vybe_compiler::primitives::object::emit_bind_method;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;
use super::{datetime_adapter, datetime_format_adapter, timespan_adapter};

const TYPE_KEY: &str = "__type";
const TIME_KEY: &str = "__time";

/// Days from `0001-01-01` (the .NET calendar origin, which is what
/// `DateOnly.DayNumber` counts from) to the Unix epoch this platform's
/// millisecond payload counts from. Verified against
/// `new DateOnly(2026,5,1).DayNumber` = 739736 = 20574 + this.
const DAYS_TO_UNIX_EPOCH: f64 = 719_162.0;

/// The total-millisecond field `timespan_adapter` mints, read here so
/// `TimeOnly.Add(TimeSpan)` and `TimeOnly.FromTimeSpan` see a duration.
const TIMESPAN_TOTAL_MS: &str = "totalmilliseconds";

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn struct_set_named_field(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

fn struct_get_named_field(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        Dest::Stack,
        line,
    );
}

fn field_from_slot(chunk: &mut Chunk, obj_slot: u16, field: &str, line: u32) {
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(obj_slot),
        &field_slot(field),
        Dest::Stack,
        line,
    );
}

fn call_import(chunks: &mut [Chunk], current: usize, module: &str, func: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, func);
    chunks[current].emit_call(idx, argc, line);
}

fn create_function_chunk(name: &str, arity: u8) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = arity;
    c
}

/// A one-argument `ToString` chunk that renders its receiver with a FIXED
/// standard pattern, through the same formatter `DateTime.ToString(fmt)` uses.
/// Deduplicated by name the way `datetime_format_adapter::format_chunk` is, so
/// a program with a hundred `DateOnly`s carries one chunk.
fn push_fixed_format_chunk(
    chunks: &mut Vec<Chunk>,
    name: &str,
    pattern: &str,
    line: u32,
) -> usize {
    if let Some(idx) = chunks.iter().position(|chunk| chunk.name == name) {
        return idx;
    }
    let mut method = create_function_chunk(name, 1);
    method.local_count = 1;
    chunks.push(method);
    let idx = chunks.len() - 1;
    chunks[idx].emit_op_u16(Op::LOCAL_GET, 0, line);
    push_const(&mut chunks[idx], Value::String(Arc::from(pattern)), line);
    datetime_format_adapter::emit_date_format(chunks, idx, line);
    chunks[idx].emit_op(Op::RETURN, line);
    idx
}

fn bind_to_string(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, method_idx: usize, line: u32) {
    emit_bind_method(
        &mut chunks[current],
        obj_slot,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString),
        method_idx,
        line,
    );
}

// ── DateOnly ────────────────────────────────────────────────────────────────

/// Wrap a millisecond timestamp as a `DateOnly`.
///
/// Stack: `[ms]` → `[dateonly_obj]`.
pub fn emit_wrap_dateonly(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let day_slot = chunks[current].alloc_scratch(2);
    let obj_slot = day_slot + 1;
    {
        // ⛔ FLOOR, not TRUNC: a date before 1970 has a negative payload and
        // truncation would round it toward the epoch, i.e. FORWARD a day.
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(MS_PER_DAY), line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_FLOOR, line);
        chunk.emit_op_u16(Op::LOCAL_SET, day_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, day_slot, line);
        push_const(chunk, Value::F64(MS_PER_DAY), line);
        chunk.emit_op(Op::F64_MUL, line);
    }
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        push_const(chunk, Value::String(Arc::from("dateonly")), line);
        struct_set_named_field(chunk, TYPE_KEY, line);
        // Both spellings, like every field `emit_wrap_ms_internal` stamps: a
        // case-insensitive frontend folds the member name, so a PascalCase-only
        // field reads `undefined` rather than erroring.
        for spelling in ["DayNumber", "daynumber"] {
            chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, day_slot, line);
            push_const(chunk, Value::F64(DAYS_TO_UNIX_EPOCH), line);
            chunk.emit_op(Op::F64_ADD, line);
            struct_set_named_field(chunk, spelling, line);
        }
    }
    let method_idx = push_fixed_format_chunk(chunks, "__dateonly_tostring", "d", line);
    bind_to_string(chunks, current, obj_slot, method_idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// `new DateOnly(year, month, day)`.
pub fn emit_dateonly_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        let day_slot = chunk.alloc_scratch(3);
        let month_slot = day_slot + 1;
        let year_slot = day_slot + 2;
        // A `calendar` argument beyond (y, m, d) sits on TOP; the Gregorian
        // calendar is the only one this constructor mints.
        for _ in 3..argc {
            chunk.emit_op(Op::DROP, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, day_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, month_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, year_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, year_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, month_slot, line);
        push_const(chunk, Value::I32(1), line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op_u16(Op::LOCAL_GET, day_slot, line);
        push_const(chunk, Value::I32(0), line);
        push_const(chunk, Value::I32(0), line);
        push_const(chunk, Value::I32(0), line);
    }
    call_import(chunks, current, "ecma:date", "UTC", 6, line);
    emit_wrap_dateonly(chunks, current, line);
}

/// `DateOnly.FromDateTime(dt)` — the calendar part of a DateTime.
pub fn emit_dateonly_from_datetime(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], TIME_KEY, line);
    emit_wrap_dateonly(chunks, current, line);
}

/// `DateOnly.FromDayNumber(n)` — `n` counts days from `0001-01-01`.
pub fn emit_dateonly_from_day_number(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(DAYS_TO_UNIX_EPOCH), line);
        chunk.emit_op(Op::F64_SUB, line);
        push_const(chunk, Value::F64(MS_PER_DAY), line);
        chunk.emit_op(Op::F64_MUL, line);
    }
    emit_wrap_dateonly(chunks, current, line);
}

/// `DateOnly.Parse(text)` — the DateTime parser, reduced to its calendar part.
pub fn emit_dateonly_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    datetime_adapter::emit_datetime_parse(chunks, current, line);
    struct_get_named_field(&mut chunks[current], TIME_KEY, line);
    emit_wrap_dateonly(chunks, current, line);
}

pub fn emit_dateonly_min_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    push_const(
        &mut chunks[current],
        Value::F64(DOTNET_DATETIME_MIN_UNIX_MS),
        line,
    );
    emit_wrap_dateonly(chunks, current, line);
}

pub fn emit_dateonly_max_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    push_const(
        &mut chunks[current],
        Value::F64(DOTNET_DATETIME_MAX_UNIX_MS),
        line,
    );
    emit_wrap_dateonly(chunks, current, line);
}

/// `d.AddDays(n)`. Stack: `[dateonly, n]` → `[dateonly]`.
pub fn emit_dateonly_add_days(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    {
        let chunk = &mut chunks[current];
        let value_slot = chunk.alloc_scratch(2);
        let date_slot = value_slot + 1;
        chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
        field_from_slot(chunk, date_slot, TIME_KEY, line);
        chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
        push_const(chunk, Value::F64(MS_PER_DAY), line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
    }
    emit_wrap_dateonly(chunks, current, line);
}

/// `d.AddMonths(n)` / `d.AddYears(n)` — the DateTime walk, re-minted as a
/// DateOnly so the type (and its `ToString`) survives the arithmetic. The
/// end-of-month clamp .NET applies (`Jan 31 + 1 month` is `Feb 28`) lives in
/// the DateTime adapter; both types answer it the same way by construction.
pub fn emit_dateonly_add_months(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    datetime_adapter::emit_datetime_add_months(chunks, current, line);
    struct_get_named_field(&mut chunks[current], TIME_KEY, line);
    emit_wrap_dateonly(chunks, current, line);
}

pub fn emit_dateonly_add_years(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    datetime_adapter::emit_datetime_add_years(chunks, current, line);
    struct_get_named_field(&mut chunks[current], TIME_KEY, line);
    emit_wrap_dateonly(chunks, current, line);
}

/// `d.ToDateTime(time)` — a DateOnly plus a TimeOnly is a DateTime.
/// `argc` counts the receiver: 1 is `ToDateTime()`, 2 adds the time, 3 adds a
/// `DateTimeKind` this surface renders UTC-relative and so ignores.
pub fn emit_dateonly_to_datetime(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        let time_slot = chunk.alloc_scratch(2);
        let date_slot = time_slot + 1;
        for _ in 3..argc {
            chunk.emit_op(Op::DROP, line);
        }
        if argc >= 2 {
            chunk.emit_op_u16(Op::LOCAL_SET, time_slot, line);
        } else {
            push_const(chunk, Value::I32(0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, time_slot, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, date_slot, line);
        field_from_slot(chunk, date_slot, TIME_KEY, line);
        if argc >= 2 {
            field_from_slot(chunk, time_slot, TIME_KEY, line);
        } else {
            chunk.emit_op_u16(Op::LOCAL_GET, time_slot, line);
        }
        chunk.emit_op(Op::F64_ADD, line);
    }
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
}

// ── TimeOnly ────────────────────────────────────────────────────────────────

/// Wrap a millisecond count as a `TimeOnly`, reduced into `[0, one day)`.
///
/// Stack: `[ms]` → `[timeonly_obj]`.
pub fn emit_wrap_timeonly(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let ms_slot = chunks[current].alloc_scratch(2);
    let obj_slot = ms_slot + 1;
    {
        // `ms - floor(ms / day) * day` — a FLOORED modulus, so a negative
        // count wraps backwards into the day the way .NET's `AddHours(-2)`
        // from `00:30` lands on `22:30` rather than on a negative time.
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
        push_const(chunk, Value::F64(MS_PER_DAY), line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_FLOOR, line);
        push_const(chunk, Value::F64(MS_PER_DAY), line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    }
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        push_const(chunk, Value::String(Arc::from("timeonly")), line);
        struct_set_named_field(chunk, TYPE_KEY, line);
        // ⛔ OVERWRITE, not add: the DateTime wrap stamped `Ticks` counting
        // from `0001-01-01`. A TimeOnly's ticks are since MIDNIGHT —
        // `new TimeOnly(1,30,0).Ticks` is 54_000_000_000, measured.
        for spelling in ["Ticks", "ticks"] {
            chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
            push_const(chunk, Value::F64(DOTNET_TICKS_PER_MS), line);
            chunk.emit_op(Op::F64_MUL, line);
            struct_set_named_field(chunk, spelling, line);
        }
    }
    let method_idx = push_fixed_format_chunk(chunks, "__timeonly_tostring", "t", line);
    bind_to_string(chunks, current, obj_slot, method_idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// `new TimeOnly(hour, minute[, second[, millisecond]])`, or
/// `new TimeOnly(ticks)`.
pub fn emit_timeonly_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        let millis_slot = chunk.alloc_scratch(4);
        let second_slot = millis_slot + 1;
        let minute_slot = millis_slot + 2;
        let hour_slot = millis_slot + 3;
        if argc == 1 {
            // The single-argument overload is a TICK count.
            push_const(chunk, Value::F64(DOTNET_TICKS_PER_MS), line);
            chunk.emit_op(Op::F64_DIV, line);
        } else {
            for _ in 4..argc {
                chunk.emit_op(Op::DROP, line);
            }
            if argc >= 4 {
                chunk.emit_op_u16(Op::LOCAL_SET, millis_slot, line);
            } else {
                push_const(chunk, Value::I32(0), line);
                chunk.emit_op_u16(Op::LOCAL_SET, millis_slot, line);
            }
            if argc >= 3 {
                chunk.emit_op_u16(Op::LOCAL_SET, second_slot, line);
            } else {
                push_const(chunk, Value::I32(0), line);
                chunk.emit_op_u16(Op::LOCAL_SET, second_slot, line);
            }
            chunk.emit_op_u16(Op::LOCAL_SET, minute_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, hour_slot, line);

            chunk.emit_op_u16(Op::LOCAL_GET, hour_slot, line);
            push_const(chunk, Value::F64(MS_PER_HOUR), line);
            chunk.emit_op(Op::F64_MUL, line);
            chunk.emit_op_u16(Op::LOCAL_GET, minute_slot, line);
            push_const(chunk, Value::F64(MS_PER_MINUTE), line);
            chunk.emit_op(Op::F64_MUL, line);
            chunk.emit_op(Op::F64_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_GET, second_slot, line);
            push_const(chunk, Value::F64(MS_PER_SECOND), line);
            chunk.emit_op(Op::F64_MUL, line);
            chunk.emit_op(Op::F64_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_GET, millis_slot, line);
            chunk.emit_op(Op::F64_ADD, line);
        }
    }
    emit_wrap_timeonly(chunks, current, line);
}

/// `TimeOnly.FromDateTime(dt)` — the clock part of a DateTime. The floored
/// modulus in the wrap is what drops the calendar part.
pub fn emit_timeonly_from_datetime(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], TIME_KEY, line);
    emit_wrap_timeonly(chunks, current, line);
}

/// `TimeOnly.FromTimeSpan(ts)`.
pub fn emit_timeonly_from_timespan(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    struct_get_named_field(&mut chunks[current], TIMESPAN_TOTAL_MS, line);
    emit_wrap_timeonly(chunks, current, line);
}

/// `TimeOnly.Parse(text)` — the DateTime parser, reduced to its clock part.
pub fn emit_timeonly_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    datetime_adapter::emit_datetime_parse(chunks, current, line);
    struct_get_named_field(&mut chunks[current], TIME_KEY, line);
    emit_wrap_timeonly(chunks, current, line);
}

pub fn emit_timeonly_min_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    push_const(&mut chunks[current], Value::F64(0.0), line);
    emit_wrap_timeonly(chunks, current, line);
}

pub fn emit_timeonly_max_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    push_const(&mut chunks[current], Value::F64(MS_PER_DAY - 1.0), line);
    emit_wrap_timeonly(chunks, current, line);
}

/// The shared body of `AddHours` / `AddMinutes`: scale the argument to
/// milliseconds and add, then let the wrap's floored modulus do the wrapping.
fn emit_timeonly_add_unit(chunks: &mut Vec<Chunk>, current: usize, unit_ms: f64, line: u32) {
    {
        let chunk = &mut chunks[current];
        let value_slot = chunk.alloc_scratch(2);
        let time_slot = value_slot + 1;
        chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, time_slot, line);
        field_from_slot(chunk, time_slot, TIME_KEY, line);
        chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
        push_const(chunk, Value::F64(unit_ms), line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
    }
    emit_wrap_timeonly(chunks, current, line);
}

pub fn emit_timeonly_add_hours(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_timeonly_add_unit(chunks, current, MS_PER_HOUR, line);
}

pub fn emit_timeonly_add_minutes(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_timeonly_add_unit(chunks, current, MS_PER_MINUTE, line);
}

/// `t.Add(TimeSpan)`.
pub fn emit_timeonly_add_timespan(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    {
        let chunk = &mut chunks[current];
        let span_slot = chunk.alloc_scratch(2);
        let time_slot = span_slot + 1;
        chunk.emit_op_u16(Op::LOCAL_SET, span_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, time_slot, line);
        field_from_slot(chunk, time_slot, TIME_KEY, line);
        field_from_slot(chunk, span_slot, TIMESPAN_TOTAL_MS, line);
        chunk.emit_op(Op::F64_ADD, line);
    }
    emit_wrap_timeonly(chunks, current, line);
}

/// `t.ToTimeSpan()` — the time elapsed since midnight.
pub fn emit_timeonly_to_timespan(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    struct_get_named_field(chunk, TIME_KEY, line);
    timespan_adapter::emit_build_timespan_from_total_ms(chunk, line);
}

/// `d.ToString()` / `t.ToString()` — the standard pattern .NET's parameterless
/// override uses for each type (`M/d/yyyy` and `h:mm tt`). The bound
/// `ToString` protocol slot answers string interpolation; this answers the
/// explicit call, and both go through the same formatter.
pub fn emit_dateonly_to_string(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    push_const(&mut chunks[current], Value::String(Arc::from("d")), line);
    datetime_format_adapter::emit_date_format(chunks, current, line);
}

pub fn emit_timeonly_to_string(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    push_const(&mut chunks[current], Value::String(Arc::from("t")), line);
    datetime_format_adapter::emit_date_format(chunks, current, line);
}
