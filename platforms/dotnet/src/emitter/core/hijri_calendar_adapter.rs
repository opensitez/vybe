//! .NET `System.Globalization.HijriCalendar` — bytecode-only.
//!
//! The type was absent from the catalog entirely, so
//! `new System.Globalization.HijriCalendar()` trapped with
//! `undefined is not callable`.
//!
//! .NET's `HijriCalendar` is the TABULAR Islamic calendar, not an observational
//! one: a 30-year cycle of exactly 10631 days, months alternating 30 and 29
//! days, and a 12th month that gains a day in a leap year. `HijriAdjustment` is
//! 0 outside Windows, so the conversion is pure arithmetic.
//!
//! ⛔ .NET's own `DaysUpToHijriYear` LOOPS over up to 29 years summing
//! `354 + isLeap(k)`. That loop has a closed form, which is what this adapter
//! emits — leaf arithmetic instead of a bytecode loop:
//!
//! ```text
//!   isLeap(y)  ⟺  (11y + 14) mod 30 < 11
//!   L(n) = #{k in 1..n : isLeap(k)} = floor((11n + 14) / 30)
//!   D(y) = 227013 + 354(y-1) + floor((11(y-1) + 14) / 30)
//! ```
//!
//! `L(n)` is exact because `(11k + 14) mod 30 < 11` holds precisely when
//! `floor((11k + 14) / 30)` steps up at `k`, so counting leap years is counting
//! those steps. Checked against .NET's own loop at the cycle boundary:
//! `D(y + 30) - D(y)` is 10631 for every `y`.
//!
//! Verified against `/usr/local/share/dotnet/dotnet` (SDK 10):
//!
//! ```text
//!   1 Jan 2026 → 1447-07-13, day-of-year 190     (this adapter: 1447-07-13, 190)
//!  20 Jan 2026 → 1447-08-02                      (this adapter: 1447-08-02)
//!   GetDaysInMonth(1447, 7) = 30, GetDaysInYear(1447) = 355, IsLeapYear = true
//!   ToDateTime(1447, 7, 12) = 2025-12-31
//! ```
//!
//! ⛔ The corpus asserts only `GetYear(...) > 1400`, which any conversion within
//! thirty years of correct satisfies. The values above are the real gate.
//!
//! `HebrewCalendar` is NOT implemented — see the module-level note in
//! `component_classes_system.rs`.

use std::sync::Arc;

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::datetime::MS_PER_DAY;
use vybe_compiler::primitives::ops::emit_dyn_gt;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;
use super::datetime_adapter;

const TYPE_KEY: &str = "__type";
const TIME_KEY: &str = "__time";

/// Days from `0001-01-01` to the Unix epoch — the offset between this
/// platform's millisecond payload and the "absolute date" .NET's calendar
/// arithmetic counts in.
const DAYS_TO_UNIX_EPOCH: f64 = 719_162.0;

/// Absolute date of `0001-01-01` in the Hijri era, i.e. the constant .NET's
/// `DaysUpToHijriYear` starts from.
const HIJRI_EPOCH_ABSOLUTE: f64 = 227_013.0;

/// Which component of a Hijri date an emitter wants.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum DatePart {
    Year,
    Month,
    Day,
    DayOfYear,
}

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

fn field_from_slot(chunk: &mut Chunk, obj_slot: u16, field: &str, line: u32) {
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(obj_slot),
        &field_slot(field),
        Dest::Stack,
        line,
    );
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// `D(year_slot)` — the absolute date of the day BEFORE that Hijri year
/// begins. Leaves one value on the stack; `tmp` is clobbered.
fn emit_days_up_to_year(chunk: &mut Chunk, year_slot: u16, tmp: u16, line: u32) {
    get(chunk, year_slot, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_SUB, line);
    set(chunk, tmp, line);

    push_const(chunk, Value::F64(HIJRI_EPOCH_ABSOLUTE), line);
    get(chunk, tmp, line);
    push_const(chunk, Value::F64(354.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);

    // + L(y - 1) = floor((11(y-1) + 14) / 30) — the leap days in the elapsed
    // years, counted as the steps of that quotient rather than by looping.
    get(chunk, tmp, line);
    push_const(chunk, Value::F64(11.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    push_const(chunk, Value::F64(14.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    push_const(chunk, Value::F64(30.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_op(Op::F64_ADD, line);
}

/// `cum(k)` — days in the first `k` months of a Hijri year, `k` in `0..=12`.
/// Months alternate 30 and 29 days, so the running total is `floor((59k+1)/2)`;
/// `cum(12)` is 354, and month 12's leap day is added separately.
fn emit_cumulative_month_days(chunk: &mut Chunk, count_slot: u16, line: u32) {
    get(chunk, count_slot, line);
    push_const(chunk, Value::F64(59.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_ADD, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
}

/// `isLeap(year_slot)` as a condition value: `(11y + 14) mod 30 < 11`.
/// Emitted as `11 > (11y + 14) - floor((11y + 14) / 30) * 30`.
fn emit_is_leap_condition(chunk: &mut Chunk, year_slot: u16, tmp: u16, line: u32) {
    get(chunk, year_slot, line);
    push_const(chunk, Value::F64(11.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    push_const(chunk, Value::F64(14.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, tmp, line);

    push_const(chunk, Value::F64(11.0), line);
    get(chunk, tmp, line);
    get(chunk, tmp, line);
    push_const(chunk, Value::F64(30.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    push_const(chunk, Value::F64(30.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    emit_dyn_gt(chunk, line);
}

/// Slots a full Gregorian→Hijri conversion needs.
struct Parts {
    n: u16,
    year: u16,
    next_year: u16,
    day_of_year: u16,
    month: u16,
    tmp: u16,
}

/// Convert the DateTime in `dt_slot` to a Hijri date, leaving `part` on the
/// stack.
fn emit_date_part_from_slot(chunk: &mut Chunk, dt_slot: u16, part: DatePart, line: u32) {
    let base = chunk.alloc_scratch(6);
    let s = Parts {
        n: base,
        year: base + 1,
        next_year: base + 2,
        day_of_year: base + 3,
        month: base + 4,
        tmp: base + 5,
    };

    // N — the absolute date, counting `0001-01-01` as 1, which is the unit
    // every constant below is expressed in.
    field_from_slot(chunk, dt_slot, TIME_KEY, line);
    push_const(chunk, Value::F64(MS_PER_DAY), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    push_const(chunk, Value::F64(DAYS_TO_UNIX_EPOCH + 1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, s.n, line);

    // The estimate .NET calls "this magic formula" — exact to within one year
    // because the cycle is exactly 10631 days.
    get(chunk, s.n, line);
    push_const(chunk, Value::F64(HIJRI_EPOCH_ABSOLUTE), line);
    chunk.emit_op(Op::F64_SUB, line);
    push_const(chunk, Value::F64(30.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    push_const(chunk, Value::F64(10631.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, s.year, line);

    // The invariant the estimate must satisfy is `D(y) < N <= D(y+1)`. Correct
    // it in each direction; one step is always enough.
    get(chunk, s.n, line);
    emit_days_up_to_year(chunk, s.year, s.tmp, line);
    emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_else(line);
    get(chunk, s.year, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_SUB, line);
    set(chunk, s.year, line);
    chunk.emit_end(line);

    get(chunk, s.year, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, s.next_year, line);
    get(chunk, s.n, line);
    emit_days_up_to_year(chunk, s.next_year, s.tmp, line);
    emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    get(chunk, s.year, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, s.year, line);
    chunk.emit_end(line);

    if part == DatePart::Year {
        get(chunk, s.year, line);
        return;
    }

    get(chunk, s.n, line);
    emit_days_up_to_year(chunk, s.year, s.tmp, line);
    chunk.emit_op(Op::F64_SUB, line);
    set(chunk, s.day_of_year, line);

    if part == DatePart::DayOfYear {
        get(chunk, s.day_of_year, line);
        return;
    }

    // The smallest `m` with `dayOfYear <= cum(m)`, inverted from
    // `cum(m) = floor((59m+1)/2)`: `m = ceil((2·doy - 1) / 59)`.
    get(chunk, s.day_of_year, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    push_const(chunk, Value::F64(57.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    push_const(chunk, Value::F64(59.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    set(chunk, s.month, line);
    // ⛔ Only the 355th day of a leap year overshoots, and it belongs to month
    // 12 — `cum(12)` is 354 because the leap day is not in the running total.
    get(chunk, s.month, line);
    push_const(chunk, Value::F64(12.0), line);
    emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::F64(12.0), line);
    set(chunk, s.month, line);
    chunk.emit_end(line);

    if part == DatePart::Month {
        get(chunk, s.month, line);
        return;
    }

    get(chunk, s.day_of_year, line);
    get(chunk, s.month, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_SUB, line);
    set(chunk, s.tmp, line);
    emit_cumulative_month_days(chunk, s.tmp, line);
    chunk.emit_op(Op::F64_SUB, line);
}

/// `new HijriCalendar()` — a stateless marker; every method is arithmetic.
pub fn emit_hijri_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        for _ in 0..argc {
            chunk.emit_op(Op::DROP, line);
        }
    }
    let idx = chunks[current].add_import("ecma:object", "new");
    chunks[current].emit_call(idx, 0, line);
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    set(chunk, obj_slot, line);
    get(chunk, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("hijricalendar")), line);
    struct_set_named_field(chunk, TYPE_KEY, line);
    get(chunk, obj_slot, line);
}

/// `cal.GetYear(dt)` and its siblings. Stack: `[cal, dt]` → `[number]`.
pub fn emit_hijri_date_part(chunks: &mut [Chunk], current: usize, part: DatePart, line: u32) {
    let chunk = &mut chunks[current];
    let dt_slot = chunk.alloc_scratch(2);
    let cal_slot = dt_slot + 1;
    set(chunk, dt_slot, line);
    set(chunk, cal_slot, line);
    emit_date_part_from_slot(chunk, dt_slot, part, line);
}

/// `cal.GetDayOfWeek(dt)` — the Gregorian weekday, which the Hijri calendar
/// shares because both count the same days.
pub fn emit_hijri_day_of_week(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let dt_slot = chunk.alloc_scratch(2);
    let cal_slot = dt_slot + 1;
    set(chunk, dt_slot, line);
    set(chunk, cal_slot, line);
    field_from_slot(chunk, dt_slot, "DayOfWeek", line);
}

/// `cal.GetEra(dt)` — `HijriEra` is 1, the only era this calendar has.
pub fn emit_hijri_era(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_const(chunk, Value::I32(1), line);
}

/// `cal.IsLeapYear(year[, era])`.
pub fn emit_hijri_is_leap_year(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let year_slot = chunk.alloc_scratch(3);
    let cal_slot = year_slot + 1;
    let tmp = year_slot + 2;
    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }
    set(chunk, year_slot, line);
    set(chunk, cal_slot, line);
    emit_is_leap_condition(chunk, year_slot, tmp, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

/// `cal.GetDaysInYear(year[, era])` — 354, or 355 in a leap year.
pub fn emit_hijri_days_in_year(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let year_slot = chunk.alloc_scratch(3);
    let cal_slot = year_slot + 1;
    let tmp = year_slot + 2;
    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }
    set(chunk, year_slot, line);
    set(chunk, cal_slot, line);
    emit_is_leap_condition(chunk, year_slot, tmp, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(355.0), line);
    chunk.emit_else(line);
    push_const(chunk, Value::F64(354.0), line);
    chunk.emit_end(line);
}

/// `cal.GetMonthsInYear(year[, era])` — the tabular calendar has no leap MONTH.
pub fn emit_hijri_months_in_year(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_const(chunk, Value::I32(12), line);
}

/// `cal.GetLeapMonth(year[, era])` — always 0: no month is intercalated.
pub fn emit_hijri_leap_month(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_const(chunk, Value::I32(0), line);
}

/// `cal.GetDaysInMonth(year, month[, era])` — `cum(m) - cum(m-1)`, plus the
/// leap day that only month 12 can gain.
pub fn emit_hijri_days_in_month(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let month_slot = chunk.alloc_scratch(5);
    let year_slot = month_slot + 1;
    let cal_slot = month_slot + 2;
    let tmp = month_slot + 3;
    let total_slot = month_slot + 4;
    for _ in 3..argc {
        chunk.emit_op(Op::DROP, line);
    }
    set(chunk, month_slot, line);
    set(chunk, year_slot, line);
    set(chunk, cal_slot, line);

    emit_cumulative_month_days(chunk, month_slot, line);
    get(chunk, month_slot, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_SUB, line);
    set(chunk, tmp, line);
    emit_cumulative_month_days(chunk, tmp, line);
    chunk.emit_op(Op::F64_SUB, line);
    set(chunk, total_slot, line);

    get(chunk, month_slot, line);
    push_const(chunk, Value::F64(11.0), line);
    emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    emit_is_leap_condition(chunk, year_slot, tmp, line);
    chunk.emit_if(line);
    get(chunk, total_slot, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, total_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    get(chunk, total_slot, line);
}

/// `cal.ToDateTime(year, month, day, hour, minute, second, millisecond)`.
pub fn emit_hijri_to_datetime(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        let millis_slot = chunk.alloc_scratch(9);
        let second_slot = millis_slot + 1;
        let minute_slot = millis_slot + 2;
        let hour_slot = millis_slot + 3;
        let day_slot = millis_slot + 4;
        let month_slot = millis_slot + 5;
        let year_slot = millis_slot + 6;
        let cal_slot = millis_slot + 7;
        let tmp = millis_slot + 8;
        // An `era` argument beyond the seven sits on TOP; there is one era.
        for _ in 8..argc {
            chunk.emit_op(Op::DROP, line);
        }
        // .NET has exactly two shapes: `(y, m, d, h, mi, s, ms)` and that plus
        // an era. Anything shorter is a midnight — there is no partial-time
        // overload to straddle, so one test covers all four slots.
        let has_time = argc >= 8;
        for slot in [millis_slot, second_slot, minute_slot, hour_slot] {
            if has_time {
                set(chunk, slot, line);
            } else {
                push_const(chunk, Value::I32(0), line);
                set(chunk, slot, line);
            }
        }
        set(chunk, day_slot, line);
        set(chunk, month_slot, line);
        set(chunk, year_slot, line);
        set(chunk, cal_slot, line);

        // absoluteDate = D(y) + cum(m - 1) + d - 1, then back to a Unix day.
        emit_days_up_to_year(chunk, year_slot, tmp, line);
        get(chunk, month_slot, line);
        push_const(chunk, Value::I32(1), line);
        chunk.emit_op(Op::F64_SUB, line);
        set(chunk, tmp, line);
        emit_cumulative_month_days(chunk, tmp, line);
        chunk.emit_op(Op::F64_ADD, line);
        get(chunk, day_slot, line);
        chunk.emit_op(Op::F64_ADD, line);
        push_const(chunk, Value::I32(1), line);
        chunk.emit_op(Op::F64_SUB, line);
        push_const(chunk, Value::F64(DAYS_TO_UNIX_EPOCH), line);
        chunk.emit_op(Op::F64_SUB, line);
        push_const(chunk, Value::F64(MS_PER_DAY), line);
        chunk.emit_op(Op::F64_MUL, line);

        for (slot, unit) in [
            (hour_slot, 3_600_000.0),
            (minute_slot, 60_000.0),
            (second_slot, 1_000.0),
        ] {
            get(chunk, slot, line);
            push_const(chunk, Value::F64(unit), line);
            chunk.emit_op(Op::F64_MUL, line);
            chunk.emit_op(Op::F64_ADD, line);
        }
        get(chunk, millis_slot, line);
        chunk.emit_op(Op::F64_ADD, line);
    }
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
}
