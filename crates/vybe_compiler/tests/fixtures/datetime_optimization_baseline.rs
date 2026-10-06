#![allow(dead_code)]
use vybe_ast::datetime::{DateTimePolicy, MonthIndexing};
use vybe_compiler::primitives::datetime::{MS_PER_DAY, month_add_clamps};
use vybe_runtime::{Chunk, Op};
pub fn emit_is_leap_year(chunk: &mut Chunk, line: u32) {
    let year = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, year, line);

    let divides = |chunk: &mut Chunk, divisor: i32| {
        chunk.emit_op_u16(Op::LOCAL_GET, year, line);
        chunk.emit_i32_const(divisor, line);
        chunk.emit_op(Op::I32_REM_S, line);
        chunk.emit_op(Op::I32_EQZ, line);
    };

    // (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
    divides(chunk, 4);
    divides(chunk, 100);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    divides(chunk, 400);
    chunk.emit_op(Op::I32_OR, line);
}

/// Days in `month` of `year`.
///
/// `indexing` says whether the caller's January is 0 or 1 — the single real
/// parameter separating the four previous implementations. Everything else
/// that differed between them (slots vs stack, f64 vs i32 result, date-object
/// vs year input) is a calling convention the caller wraps.
///
/// Stack: `[year, month]` → `[days]`.
pub fn emit_days_in_month(chunk: &mut Chunk, indexing: MonthIndexing, line: u32) {
    if indexing == MonthIndexing::ZeroBased {
        // Normalise to 1-based once, so the table below reads like a calendar.
        chunk.emit_f64_const(1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
    }
    let month = chunk.alloc_scratch(1);
    let year = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, month, line);
    chunk.emit_op_u16(Op::LOCAL_SET, year, line);

    // February is the only month needing the year; the rest are a lookup.
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_f64_const(2.0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, year, line);
    emit_is_leap_year(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(29.0, line);
    chunk.emit_else(line);
    chunk.emit_f64_const(28.0, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    // 30 days hath September (9), April (4), June (6) and November (11).
    for short_month in [4.0, 6.0, 9.0, 11.0] {
        chunk.emit_op_u16(Op::LOCAL_GET, month, line);
        chunk.emit_f64_const(short_month, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if_value(line);
        chunk.emit_f64_const(30.0, line);
        chunk.emit_else(line);
    }
    chunk.emit_f64_const(31.0, line);
    for _ in 0..4 {
        chunk.emit_end(line);
    }
    chunk.emit_end(line);
}

pub fn emit_add_months(chunk: &mut Chunk, policy: DateTimePolicy, line: u32) {
    let n = chunk.alloc_scratch(1);
    let ms = chunk.alloc_scratch(1);
    let y = chunk.alloc_scratch(1);
    let mo = chunk.alloc_scratch(1);
    let day = chunk.alloc_scratch(1);
    let time_of_day = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ms, line);

    // Decompose. The host owns the calendar, so ask it rather than reimplement
    // civil-from-days here.
    let getter = |chunk: &mut Chunk, name: &str, slot: u16, line: u32| {
        chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
        let idx = chunk.add_import("ecma:date", name);
        chunk.emit_call(idx, 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    };
    getter(chunk, "getUTCFullYear", y, line);
    getter(chunk, "getUTCMonth", mo, line);
    getter(chunk, "getUTCDate", day, line);

    // Keep the wall-clock time of day: ms - startOfDayMs.
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms, line);
    chunk.emit_f64_const(MS_PER_DAY, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_f64_const(MS_PER_DAY, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, time_of_day, line);

    // total = month + n, then y += floor(total/12), month = total mod 12.
    // `Date.UTC` normalises out-of-range months by itself, so under Overflow
    // the shift can simply be handed to it — that IS the overflow behaviour.
    chunk.emit_op_u16(Op::LOCAL_GET, mo, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, mo, line);

    if month_add_clamps(policy) {
        // Clamp: pin the day to the target month's length before rebuilding,
        // so surplus days cannot roll forward.
        let normalized_y = chunk.alloc_scratch(1);
        let normalized_m = chunk.alloc_scratch(1);
        // y + floor(mo / 12)
        chunk.emit_op_u16(Op::LOCAL_GET, y, line);
        chunk.emit_op_u16(Op::LOCAL_GET, mo, line);
        chunk.emit_f64_const(12.0, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_FLOOR, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, normalized_y, line);
        // mo - 12 * floor(mo / 12)  → 0..11, correct for negative n too.
        chunk.emit_op_u16(Op::LOCAL_GET, mo, line);
        chunk.emit_op_u16(Op::LOCAL_GET, mo, line);
        chunk.emit_f64_const(12.0, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_FLOOR, line);
        chunk.emit_f64_const(12.0, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op_u16(Op::LOCAL_SET, normalized_m, line);

        // day = min(day, days_in_month(y', m'))
        let limit = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_GET, normalized_y, line);
        chunk.emit_op_u16(Op::LOCAL_GET, normalized_m, line);
        emit_days_in_month(chunk, MonthIndexing::ZeroBased, line);
        chunk.emit_op_u16(Op::LOCAL_SET, limit, line);
        chunk.emit_op_u16(Op::LOCAL_GET, day, line);
        chunk.emit_op_u16(Op::LOCAL_GET, limit, line);
        chunk.emit_op(Op::F64_GT, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, limit, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, day, line);
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_SET, day, line);

        chunk.emit_op_u16(Op::LOCAL_GET, normalized_y, line);
        chunk.emit_op_u16(Op::LOCAL_GET, normalized_m, line);
    } else {
        chunk.emit_op_u16(Op::LOCAL_GET, y, line);
        chunk.emit_op_u16(Op::LOCAL_GET, mo, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, day, line);
    let utc = chunk.add_import("ecma:date", "UTC");
    chunk.emit_call(utc, 3, line);
    chunk.emit_op_u16(Op::LOCAL_GET, time_of_day, line);
    chunk.emit_op(Op::F64_ADD, line);
}
