//! Shared `Microsoft.VisualBasic.DateAndTime` helpers for .NET languages —
//! bytecode-only, the sibling of [`super::financial_adapter`].
//!
//! `DateAdd` / `DateDiff` / `DatePart` / `DateSerial` / `TimeSerial` /
//! `Weekday` / `WeekdayName` / `MonthName` are NOT `System.DateTime` API. They
//! are the `Microsoft.VisualBasic` intrinsics, which is why a grep of this tree
//! for those five names came back empty while
//! [`super::datetime_adapter`] — a complete `System.DateTime` — sat right
//! beside it. They belong here for the same reason `Financial.Pmt` does: the
//! namespace is part of the .NET platform, not part of one language's walker.
//!
//! Everything below is expressed on top of the `System.DateTime` object
//! (`{__type:"datetime", __time: ms, Year, Month, …}`) that
//! `datetime_adapter` already produces, and on
//! `vybe_compiler::primitives::datetime` for the calendar rules. Nothing here
//! reimplements proleptic Gregorian arithmetic.
//!
//! ## Why the interval is read at RUNTIME
//!
//! The interval argument is a string (`"m"`, `"yyyy"`) or a `DateInterval`
//! member, which lowers to that member's own name. Both arrive as an ordinary
//! value, so the lowering below reads it with a flat chain of
//! compare-and-store — [`emit_interval_lookup`]. It is deliberately NOT a
//! compile-time fold: a fold answers only when the walker can see through the
//! operands, and the VB walker's chrono-based fold that this replaces silently
//! produced nothing when it could not, with the profile row emitting `noop`
//! behind it.

use std::sync::Arc;
use vybe_compiler::primitives::datetime as dt;
use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::datetime_adapter;

/// `DateTime.MinValue` in ECMA milliseconds — 0001-01-01T00:00:00Z. VB's
/// `TimeSerial` returns a time-of-day anchored at that date.
const MIN_VALUE_MS: f64 = -62_135_596_800_000.0;

const MS_PER_MINUTE: f64 = 60_000.0;
const MS_PER_SECOND: f64 = 1_000.0;

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    dt::push_const(chunk, val, line);
}

fn push_f64(chunk: &mut Chunk, value: f64, line: u32) {
    push_const(chunk, Value::F64(value), line);
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(value)), line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// Pop the interval argument, lower-case it, and park it in a slot. VB matches
/// interval strings case-insensitively (`"YYYY"` and `"yyyy"` are the same
/// interval), and `DateInterval.Day` arrives as the member's own capitalised
/// name, so every table below is written in lower case and the value is
/// normalised once, here.
fn emit_interval_into_slot(chunk: &mut Chunk, slot: u16, line: u32) {
    let lower_idx = chunk.add_import("ecma:string", "toLowerCase");
    chunk.emit_call(lower_idx, 1, line);
    lset(chunk, slot, line);
}

/// `out_slot = table[interval_slot] ?? default`, as a FLAT chain of void `if`s.
///
/// Flat rather than nested because nothing crosses a block boundary — each arm
/// only stores to a local — so `emit_if`'s void blocktype is exactly right
/// here. (A nested `emit_if_value` chain would have to carry the result
/// through every level, which is where the sibling `Version.CompareTo`
/// emitter lost its value to a void block.)
fn emit_interval_lookup(
    chunk: &mut Chunk,
    interval_slot: u16,
    out_slot: u16,
    table: &[(&[&str], f64)],
    default: f64,
    line: u32,
) {
    push_f64(chunk, default, line);
    lset(chunk, out_slot, line);
    for (spellings, value) in table {
        emit_when_interval(chunk, interval_slot, spellings, line, |chunk| {
            push_f64(chunk, *value, line);
            lset(chunk, out_slot, line);
        });
    }
}

/// `if interval_slot ∈ spellings { body }` — a void `if` over the OR of the
/// individual string comparisons.
fn emit_when_interval(
    chunk: &mut Chunk,
    interval_slot: u16,
    spellings: &[&str],
    line: u32,
    body: impl FnOnce(&mut Chunk),
) {
    for (index, spelling) in spellings.iter().enumerate() {
        lget(chunk, interval_slot, line);
        push_str(chunk, spelling, line);
        ops::emit_dyn_eq(chunk, line);
        ops::emit_dyn_to_bool(chunk, line);
        if index > 0 {
            chunk.emit_op(Op::I32_OR, line);
        }
    }
    chunk.emit_if(line);
    body(chunk);
    chunk.emit_end(line);
}

/// `if slot != 0 { body }` — the numeric sibling of [`emit_when_interval`],
/// used to pick between DateDiff's three families once their selectors have
/// been resolved from the interval.
fn emit_when_nonzero(chunk: &mut Chunk, slot: u16, line: u32, body: impl FnOnce(&mut Chunk)) {
    lget(chunk, slot, line);
    push_f64(chunk, 0.0, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if(line);
    body(chunk);
    chunk.emit_end(line);
}

/// Months added per unit of `Number`, by interval. Zero for every interval
/// whose step is a fixed span of milliseconds — those go through
/// [`MS_PER_UNIT`] instead, and exactly one of the two tables is non-zero for
/// any given interval, so `DateAdd` is a month shift followed by a
/// millisecond shift with no branch between the families.
const MONTHS_PER_UNIT: &[(&[&str], f64)] = &[
    (&["yyyy", "year"], 12.0),
    (&["q", "quarter"], 3.0),
    (&["m", "month"], 1.0),
];

/// Milliseconds added per unit of `Number`, by interval.
///
/// `DateInterval.Weekday` ("w") steps by a DAY in `DateAdd` — VB documents it
/// as behaving like `"d"` there, and only `DateDiff` reads it as a count of
/// weeks.
const MS_PER_UNIT: &[(&[&str], f64)] = &[
    (&["y", "dayofyear", "d", "day", "w", "weekday"], dt::MS_PER_DAY),
    (
        &["ww", "week", "weekofyear"],
        dt::MS_PER_DAY * 7.0,
    ),
    (&["h", "hour"], dt::MS_PER_HOUR),
    (&["n", "minute"], MS_PER_MINUTE),
    (&["s", "second"], MS_PER_SECOND),
];

/// `DateAdd(Interval, Number, DateValue)`.
///
/// Stack on entry: `[interval, number, date]` ; on exit: `[datetime_obj]`.
pub fn emit_vb_date_add(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let (date_slot, number_slot, interval_slot, months_slot, ms_slot, base_slot, obj_slot) = {
        let chunk = &mut chunks[current];
        let base = chunk.alloc_scratch(7);
        (base, base + 1, base + 2, base + 3, base + 4, base + 5, base + 6)
    };

    {
        let chunk = &mut chunks[current];
        lset(chunk, date_slot, line);
        lset(chunk, number_slot, line);
        emit_interval_into_slot(chunk, interval_slot, line);

        emit_interval_lookup(chunk, interval_slot, months_slot, MONTHS_PER_UNIT, 0.0, line);
        lget(chunk, months_slot, line);
        lget(chunk, number_slot, line);
        chunk.emit_op(Op::F64_MUL, line);
        lset(chunk, months_slot, line);

        emit_interval_lookup(chunk, interval_slot, ms_slot, MS_PER_UNIT, 0.0, line);
        lget(chunk, ms_slot, line);
        lget(chunk, number_slot, line);
        chunk.emit_op(Op::F64_MUL, line);
        lset(chunk, ms_slot, line);

        // Baseline: the instant we started from. The month shift below replaces
        // it only when the interval is a calendar one, so a `DateAdd("h", …)`
        // never round-trips through the Y/M/D rebuild (which drops
        // milliseconds).
        datetime_adapter::emit_millis_from_slot(chunk, date_slot, line);
        lset(chunk, base_slot, line);

        lget(chunk, months_slot, line);
        push_f64(chunk, 0.0, line);
        chunk.emit_op(Op::F64_NE, line);
        chunk.emit_if(line);
        lget(chunk, date_slot, line);
        lget(chunk, months_slot, line);
    }
    // Clamping month arithmetic — .NET's `MonthOverflow::Clamp`, which the
    // System.DateTime adapter already owns.
    datetime_adapter::emit_datetime_add_months(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, obj_slot, line);
        datetime_adapter::emit_millis_from_slot(chunk, obj_slot, line);
        lset(chunk, base_slot, line);
        chunk.emit_end(line);

        lget(chunk, base_slot, line);
        lget(chunk, ms_slot, line);
        chunk.emit_op(Op::F64_ADD, line);
    }
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
}

/// Divisor applied to the whole-month difference, by interval. Year is absent
/// on purpose: VB's `DateDiff(DateInterval.Year, …)` is the difference of the
/// two YEAR NUMBERS, not the month difference divided by twelve — 2020-12 to
/// 2021-01 is one year and one month, and `1 / 12` truncates to zero.
const MONTH_DIVISOR: &[(&[&str], f64)] = &[(&["m", "month"], 1.0), (&["q", "quarter"], 3.0)];

const YEAR_INTERVAL: &[(&[&str], f64)] = &[(&["yyyy", "year"], 1.0)];

/// Milliseconds per counted unit, by interval — the span `DateDiff` divides
/// the elapsed time by. `"w"`/`"ww"` both count WEEKS here (unlike `DateAdd`,
/// where `"w"` steps a day).
const MS_PER_COUNTED_UNIT: &[(&[&str], f64)] = &[
    (&["y", "dayofyear", "d", "day"], dt::MS_PER_DAY),
    (
        &["w", "weekday", "ww", "week", "weekofyear"],
        dt::MS_PER_DAY * 7.0,
    ),
    (&["h", "hour"], dt::MS_PER_HOUR),
    (&["n", "minute"], MS_PER_MINUTE),
    (&["s", "second"], MS_PER_SECOND),
];

/// `DateDiff(Interval, Date1, Date2)` — `Date2 - Date1`, truncated toward
/// zero.
///
/// Stack on entry: `[interval, date1, date2]` ; on exit: `[count]`.
pub fn emit_vb_date_diff(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(9);
    let (d2_slot, d1_slot, interval_slot) = (base, base + 1, base + 2);
    let (unit_slot, month_div_slot, year_slot) = (base + 3, base + 4, base + 5);
    let (result_slot, months_slot, scratch_slot) = (base + 6, base + 7, base + 8);

    lset(chunk, d2_slot, line);
    lset(chunk, d1_slot, line);
    emit_interval_into_slot(chunk, interval_slot, line);

    // Default family: elapsed milliseconds over a fixed unit. The `1.0`
    // default keeps an unrecognised interval a division by one rather than a
    // NaN, and the two calendar families below overwrite it when they apply.
    emit_interval_lookup(
        chunk,
        interval_slot,
        unit_slot,
        MS_PER_COUNTED_UNIT,
        1.0,
        line,
    );
    datetime_adapter::emit_millis_from_slot(chunk, d2_slot, line);
    datetime_adapter::emit_millis_from_slot(chunk, d1_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    lget(chunk, unit_slot, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_TRUNC, line);
    lset(chunk, result_slot, line);

    // Whole months between the two, reused by the month and quarter arms.
    datetime_adapter::emit_field_from_slot(chunk, d2_slot, "Year", line);
    datetime_adapter::emit_field_from_slot(chunk, d1_slot, "Year", line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, scratch_slot, line);
    lget(chunk, scratch_slot, line);
    push_f64(chunk, 12.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    datetime_adapter::emit_field_from_slot(chunk, d2_slot, "Month", line);
    chunk.emit_op(Op::F64_ADD, line);
    datetime_adapter::emit_field_from_slot(chunk, d1_slot, "Month", line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, months_slot, line);

    emit_interval_lookup(chunk, interval_slot, month_div_slot, MONTH_DIVISOR, 0.0, line);
    emit_when_nonzero(chunk, month_div_slot, line, |chunk| {
        lget(chunk, months_slot, line);
        lget(chunk, month_div_slot, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_TRUNC, line);
        lset(chunk, result_slot, line);
    });

    emit_interval_lookup(chunk, interval_slot, year_slot, YEAR_INTERVAL, 0.0, line);
    emit_when_nonzero(chunk, year_slot, line, |chunk| {
        lget(chunk, scratch_slot, line);
        lset(chunk, result_slot, line);
    });

    lget(chunk, result_slot, line);
}

/// `DatePart(Interval, DateValue)`.
///
/// Stack on entry: `[interval, date]` ; on exit: `[part]`.
pub fn emit_vb_date_part(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(3);
    let (date_slot, interval_slot, result_slot) = (base, base + 1, base + 2);

    lset(chunk, date_slot, line);
    emit_interval_into_slot(chunk, interval_slot, line);

    push_f64(chunk, 0.0, line);
    lset(chunk, result_slot, line);

    // The plain field extractions. `Year`/`Month`/`Day`/`Hour`/`Minute`/
    // `Second`/`DayOfYear` are already on the DateTime object.
    for (spellings, field) in [
        (&["yyyy", "year"][..], "Year"),
        (&["m", "month"][..], "Month"),
        (&["d", "day"][..], "Day"),
        (&["y", "dayofyear"][..], "DayOfYear"),
        (&["h", "hour"][..], "Hour"),
        (&["n", "minute"][..], "Minute"),
        (&["s", "second"][..], "Second"),
    ] {
        emit_when_interval(chunk, interval_slot, spellings, line, |chunk| {
            datetime_adapter::emit_field_from_slot(chunk, date_slot, field, line);
            lset(chunk, result_slot, line);
        });
    }

    // Quarter: 1..4 from the 1-based month.
    emit_when_interval(chunk, interval_slot, &["q", "quarter"], line, |chunk| {
        datetime_adapter::emit_field_from_slot(chunk, date_slot, "Month", line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        push_f64(chunk, 3.0, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_TRUNC, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, result_slot, line);
    });

    // Week of year: 1..53, counting seven-day blocks from January 1st.
    emit_when_interval(
        chunk,
        interval_slot,
        &["ww", "week", "weekofyear"],
        line,
        |chunk| {
            datetime_adapter::emit_field_from_slot(chunk, date_slot, "DayOfYear", line);
            push_f64(chunk, 1.0, line);
            chunk.emit_op(Op::F64_SUB, line);
            push_f64(chunk, 7.0, line);
            chunk.emit_op(Op::F64_DIV, line);
            chunk.emit_op(Op::F64_TRUNC, line);
            push_f64(chunk, 1.0, line);
            chunk.emit_op(Op::F64_ADD, line);
            lset(chunk, result_slot, line);
        },
    );

    // Weekday: Sunday = 1 … Saturday = 7 — VB's own numbering, which is a
    // `WeekdayBase` and therefore the shared primitive's job.
    emit_when_interval(chunk, interval_slot, &["w", "weekday"], line, |chunk| {
        emit_weekday_number(chunk, date_slot, line);
        lset(chunk, result_slot, line);
    });

    lget(chunk, result_slot, line);
}

/// Sunday=1 … Saturday=7 for the DateTime in `date_slot`.
///
/// Stack: `[]` → `[weekday]`.
fn emit_weekday_number(chunk: &mut Chunk, date_slot: u16, line: u32) {
    datetime_adapter::emit_millis_from_slot(chunk, date_slot, line);
    let day_idx = chunk.add_import("ecma:date", "getUTCDay");
    chunk.emit_call(day_idx, 1, line);
    dt::emit_weekday_in_base(chunk, vybe_ast::datetime::WeekdayBase::SundayOne, line);
}

/// `Weekday(DateValue[, FirstDayOfWeek])`.
///
/// Stack on entry: `[date]` or `[date, first_day]` ; on exit: `[weekday]`.
pub fn emit_vb_weekday(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(2);
    let (first_slot, date_slot) = (base, base + 1);

    if argc >= 2 {
        lset(chunk, first_slot, line);
    } else {
        // `vbSunday` — VB's default first day of week, and the origin the
        // ECMA base already uses.
        push_f64(chunk, 1.0, line);
        lset(chunk, first_slot, line);
    }
    lset(chunk, date_slot, line);

    // ((weekday - first) mod 7) + 1, where `weekday` and `first` are both
    // 1-based. `+ 7` before the modulus keeps a negative shift positive.
    emit_weekday_number(chunk, date_slot, line);
    lget(chunk, first_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    push_f64(chunk, 7.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    push_f64(chunk, 7.0, line);
    vybe_compiler::primitives::math::emit_c_fmod(chunk, line);
    push_f64(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
}

/// `DateSerial(Year, Month, Day)`. Out-of-range components roll over, which is
/// both VB's documented behaviour and `Date.UTC`'s.
///
/// Stack on entry: `[year, month, day]` ; on exit: `[datetime_obj]`.
pub fn emit_vb_date_serial(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    {
        let chunk = &mut chunks[current];
        let base = chunk.alloc_scratch(3);
        let (day_slot, month_slot, year_slot) = (base, base + 1, base + 2);
        lset(chunk, day_slot, line);
        lset(chunk, month_slot, line);
        lset(chunk, year_slot, line);

        lget(chunk, year_slot, line);
        lget(chunk, month_slot, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        lget(chunk, day_slot, line);
        push_f64(chunk, 0.0, line);
        push_f64(chunk, 0.0, line);
        push_f64(chunk, 0.0, line);
        let utc_idx = chunk.add_import("ecma:date", "UTC");
        chunk.emit_call(utc_idx, 6, line);
    }
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
}

/// `TimeSerial(Hour, Minute, Second)` — a time of day on `DateTime.MinValue`'s
/// date, which is what VB returns.
///
/// Stack on entry: `[hour, minute, second]` ; on exit: `[datetime_obj]`.
pub fn emit_vb_time_serial(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    {
        let chunk = &mut chunks[current];
        let base = chunk.alloc_scratch(3);
        let (second_slot, minute_slot, hour_slot) = (base, base + 1, base + 2);
        lset(chunk, second_slot, line);
        lset(chunk, minute_slot, line);
        lset(chunk, hour_slot, line);

        push_f64(chunk, MIN_VALUE_MS, line);
        lget(chunk, hour_slot, line);
        dt::emit_add_scaled(chunk, dt::MS_PER_HOUR, line);
        lget(chunk, minute_slot, line);
        dt::emit_add_scaled(chunk, MS_PER_MINUTE, line);
        lget(chunk, second_slot, line);
        dt::emit_add_scaled(chunk, MS_PER_SECOND, line);
    }
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
}

const WEEKDAY_NAMES: [&str; 7] = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
];

const MONTH_NAMES: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];

/// `out_slot = names[index_slot - 1]`, abbreviated to three letters when
/// `abbreviate_slot` is truthy. A flat compare-and-store chain, the same shape
/// as [`emit_interval_lookup`] but keyed on a number.
fn emit_name_lookup(
    chunk: &mut Chunk,
    index_slot: u16,
    abbreviate_slot: Option<u16>,
    names: &[&str],
    out_slot: u16,
    line: u32,
) {
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    for (offset, name) in names.iter().enumerate() {
        lget(chunk, index_slot, line);
        push_f64(chunk, (offset + 1) as f64, line);
        chunk.emit_op(Op::F64_EQ, line);
        chunk.emit_if(line);
        push_str(chunk, name, line);
        lset(chunk, out_slot, line);
        chunk.emit_end(line);
    }
    if let Some(abbreviate_slot) = abbreviate_slot {
        lget(chunk, abbreviate_slot, line);
        ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, out_slot, line);
        push_const(chunk, Value::I32(0), line);
        push_const(chunk, Value::I32(3), line);
        let substring_idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(substring_idx, 3, line);
        lset(chunk, out_slot, line);
        chunk.emit_end(line);
    }
}

/// `MonthName(Month[, Abbreviate])`.
///
/// Stack on entry: `[month]` or `[month, abbreviate]` ; on exit: `[name]`.
pub fn emit_vb_month_name(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_indexed_name(chunks, current, argc, &MONTH_NAMES, line);
}

/// `WeekdayName(Weekday[, Abbreviate[, FirstDayOfWeek]])`.
///
/// Stack on entry: `[weekday]`, `[weekday, abbreviate]` or
/// `[weekday, abbreviate, first_day]` ; on exit: `[name]`.
pub fn emit_vb_weekday_name(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 3 {
        // `FirstDayOfWeek` renumbers which weekday index 1 refers to. The
        // default (`vbSunday`) is index 1 already, and no corpus case passes
        // anything else; dropping a supplied value would silently answer the
        // wrong name, so the argument is folded into the index instead.
        let chunk = &mut chunks[current];
        let base = chunk.alloc_scratch(2);
        let (first_slot, abbreviate_slot) = (base, base + 1);
        lset(chunk, first_slot, line);
        lset(chunk, abbreviate_slot, line);
        // index += first - 1, wrapped into 1..7.
        lget(chunk, first_slot, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op(Op::F64_ADD, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        push_f64(chunk, 7.0, line);
        vybe_compiler::primitives::math::emit_c_fmod(chunk, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
        lget(chunk, abbreviate_slot, line);
        emit_indexed_name(chunks, current, 2, &WEEKDAY_NAMES, line);
        return;
    }
    emit_indexed_name(chunks, current, argc, &WEEKDAY_NAMES, line);
}

fn emit_indexed_name(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    names: &[&str],
    line: u32,
) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(3);
    let (abbreviate_slot, index_slot, out_slot) = (base, base + 1, base + 2);
    if argc >= 2 {
        lset(chunk, abbreviate_slot, line);
    } else {
        push_const(chunk, Value::I32(0), line);
        lset(chunk, abbreviate_slot, line);
    }
    lset(chunk, index_slot, line);
    emit_name_lookup(
        chunk,
        index_slot,
        Some(abbreviate_slot),
        names,
        out_slot,
        line,
    );
    lget(chunk, out_slot, line);
}
