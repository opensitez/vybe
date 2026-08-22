//! Shared calendar + time-zone emitters.
//!
//! Dates were the conspicuous hole in `primitives/`: 50-odd modules for arrays,
//! collections, dicts, maths, JSON, errors, generators — and nothing for time,
//! while ~8,400 lines of date codegen sat in six per-language adapters. The
//! consequence was four independent implementations of identical proleptic
//! Gregorian arithmetic (`emit_days_in_month` in PHP and Python,
//! `emit_time_is_leap_year` in Java, `emit_datetime_is_leap_year` in .NET), and
//! duplicated plumbing beneath them (`struct_get`, `push_const`, `call_import`).
//!
//! Split of responsibilities, so this module stays the right size:
//!
//! | Layer | Owns | Where |
//! |---|---|---|
//! | Data + locale semantics | tzdb offsets, DST, CLDR formatting | `platforms/ecma` host |
//! | Calendar arithmetic + host glue | this module | `primitives/datetime.rs` |
//! | Language spelling | format letters, `DateInterval`, `strtotime` | the adapter |
//!
//! Time zones live HERE rather than in a `primitives/timezone.rs`: at the
//! emitter level a zone operation is a handful of host calls plus unit
//! conversion, so a separate module would be nearly empty and would invite an
//! arbitrary question — does `setTimezone` belong to dates or to zones? The
//! substantial timezone code is data, and data cannot live in an emitter.
//!
//! Everything that varies by language is a [`DateTimePolicy`] parameter, never
//! a language check — see that type for why each knob exists.

use vybe_ast::datetime::{
    DateTimePolicy, EpochPrecision, MonthIndexing, MonthOverflow, WeekdayBase,
};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

/// The host module every time-zone answer comes from. Named once so no adapter
/// re-spells it and drifts.
pub const TZ_MODULE: &str = "ecma:intl/timezone";

/// Scale a language-native epoch value to ECMA milliseconds, the unit all
/// shared arithmetic works in.
///
/// Stack: `[epoch]` → `[ms]`.
pub fn emit_epoch_to_millis(chunk: &mut Chunk, precision: EpochPrecision, line: u32) {
    let factor = match precision {
        EpochPrecision::Seconds => 1_000.0,
        EpochPrecision::Millis => return,
        EpochPrecision::Micros => 0.001,
        EpochPrecision::Nanos => 0.000_001,
    };
    chunk.emit_f64_const(factor, line);
    chunk.emit_op(Op::F64_MUL, line);
}

/// Inverse of [`emit_epoch_to_millis`] — back to the language's own unit.
///
/// Stack: `[ms]` → `[epoch]`.
pub fn emit_millis_to_epoch(chunk: &mut Chunk, precision: EpochPrecision, line: u32) {
    let factor = match precision {
        EpochPrecision::Seconds => 0.001,
        EpochPrecision::Millis => return,
        EpochPrecision::Micros => 1_000.0,
        EpochPrecision::Nanos => 1_000_000.0,
    };
    chunk.emit_f64_const(factor, line);
    chunk.emit_op(Op::F64_MUL, line);
}

/// Offset of `zone` at `ms`, in SECONDS EAST of UTC — tzdb's sign convention,
/// and PHP `DateTimeZone::getOffset`'s.
///
/// Deliberately NOT JavaScript's convention: `Date.prototype.getTimezoneOffset`
/// reports minutes WEST, i.e. opposite sign and different unit. Converting is
/// left to [`emit_offset_to_js_minutes`] so the discrepancy is visible at the
/// call site instead of being silently absorbed here.
///
/// Stack: `[zone, ms]` → `[seconds_east]`.
pub fn emit_zone_offset_seconds(chunk: &mut Chunk, line: u32) {
    let import = chunk.add_import(TZ_MODULE, "offset");
    chunk.emit_call(import, 2, line);
}

/// Convert seconds-east to JavaScript's minutes-west.
///
/// Stack: `[seconds_east]` → `[minutes_west]`.
pub fn emit_offset_to_js_minutes(chunk: &mut Chunk, line: u32) {
    chunk.emit_f64_const(-60.0, line);
    chunk.emit_op(Op::F64_DIV, line);
}

/// The host environment's zone identifier — ECMA-262 `SystemTimeZoneIdentifier`.
/// One clock for every language, so a zone set from one is observed by all.
///
/// Stack: `[]` → `[zone]`.
pub fn emit_system_zone(chunk: &mut Chunk, line: u32) {
    let import = chunk.add_import(TZ_MODULE, "systemIdentifier");
    chunk.emit_call(import, 0, line);
}

/// Set the host environment's zone; pushes a bool for whether the identifier
/// was recognised. Backs PHP `date_default_timezone_set`, Java
/// `TimeZone.setDefault`, .NET `TimeZoneInfo`.
///
/// Stack: `[zone]` → `[ok]`.
pub fn emit_set_system_zone(chunk: &mut Chunk, line: u32) {
    let import = chunk.add_import(TZ_MODULE, "setSystemIdentifier");
    chunk.emit_call(import, 1, line);
}

/// Whether daylight saving is in effect for `zone` at `ms`.
///
/// Stack: `[zone, ms]` → `[bool]`.
pub fn emit_zone_is_dst(chunk: &mut Chunk, line: u32) {
    let import = chunk.add_import(TZ_MODULE, "isDst");
    chunk.emit_call(import, 2, line);
}

/// Zone abbreviation in effect at an instant (`EST`, `BST`, …).
///
/// Stack: `[zone, ms]` → `[string]`.
pub fn emit_zone_abbreviation(chunk: &mut Chunk, line: u32) {
    let import = chunk.add_import(TZ_MODULE, "abbreviation");
    chunk.emit_call(import, 2, line);
}

/// Canonical (tzdb-cased, primary) identifier, or null when unknown. Backs
/// ECMA-402 `CanonicalizeTimeZoneName` and every language's zone validation.
///
/// Stack: `[zone]` → `[canonical|null]`.
pub fn emit_zone_canonicalize(chunk: &mut Chunk, line: u32) {
    let import = chunk.add_import(TZ_MODULE, "canonicalize");
    chunk.emit_call(import, 1, line);
}

/// Is `year` a leap year — proleptic Gregorian, the ONE rule every language
/// agrees on, which is why it takes no policy and why four copies of it was
/// four chances to get it wrong.
///
/// Integer ops, not float: this mirrors Python's `emit_cal_isleap`, which was
/// the best of the four existing implementations. PHP's used a Feb-29 rollover
/// through `ecma:date.UTC` — correct, but it needs a date object and a host
/// call to answer a question about an integer.
///
/// Stack: `[year_i32]` → `[i32 0|1]`.
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
    super::ops::emit_dyn_eq(chunk, line);
    super::ops::emit_dyn_to_bool(chunk, line);
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
        super::ops::emit_dyn_eq(chunk, line);
        super::ops::emit_dyn_to_bool(chunk, line);
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

/// Normalise a weekday from the ECMA base (Sunday = 0) into the language's own
/// numbering. `emit_*` callers push a `getDay()` result and get back whatever
/// their surface promises.
///
/// Stack: `[sunday_zero]` → `[weekday]`.
pub fn emit_weekday_in_base(chunk: &mut Chunk, base: WeekdayBase, line: u32) {
    match base {
        WeekdayBase::SundayZero => {}
        WeekdayBase::MondayOne => {
            // ((d + 6) % 7) + 1 → Monday=1 … Sunday=7
            chunk.emit_f64_const(6.0, line);
            chunk.emit_op(Op::F64_ADD, line);
            chunk.emit_f64_const(7.0, line);
            super::math::emit_c_fmod(chunk, line);
            chunk.emit_f64_const(1.0, line);
            chunk.emit_op(Op::F64_ADD, line);
        }
        WeekdayBase::MondayZero => {
            // (d + 6) % 7 → Monday=0 … Sunday=6
            chunk.emit_f64_const(6.0, line);
            chunk.emit_op(Op::F64_ADD, line);
            chunk.emit_f64_const(7.0, line);
            super::math::emit_c_fmod(chunk, line);
        }
        WeekdayBase::SundayOne => {
            // d + 1 → Sunday=1 … Saturday=7. No modulus: the ECMA base already
            // starts at Sunday, so only the origin moves.
            chunk.emit_f64_const(1.0, line);
            chunk.emit_op(Op::F64_ADD, line);
        }
    }
}

/// Whether month arithmetic must clamp the day-of-month after shifting.
///
/// Exposed as a predicate rather than baked in because the CLAMP branch needs
/// an extra `days_in_month` probe that the OVERFLOW branch must not emit —
/// PHP/JS want `Jan 31 + 1 month` to become `Mar 2`, Java/.NET/Python want
/// `Feb 29`.
pub fn month_add_clamps(policy: DateTimePolicy) -> bool {
    policy.month_overflow == MonthOverflow::Clamp
}

/// Push a constant as the WASM instruction that expresses it — **the one
/// constant-encoding policy for the whole compiler.** `Compiler::emit_const`
/// (`control_flow.rs`) delegates here, so the method used by the core
/// expression compiler and the free function used by 1432 adapter call sites
/// cannot drift apart.
///
/// The numeric variants are core `*.const`; `null` is `ref.null`; boolean,
/// string and bigint are the js-primitive-builtins / js-string-builtins
/// imports (`wasm:js-boolean.fromI32`, `wasm:string-constants`,
/// `wasm:js-bigint.fromI64`). None of it goes through the constant pool, and
/// none of it is a custom opcode.
///
/// Anything with no WASM encoding panics rather than pooling it. The pool push
/// this replaced looked safe but wasn't: the wasm writer's fallback lowers an
/// unrecognised pool constant to `ref.null extern`, so an `Object` or `Symbol`
/// arriving here used to become a silent null in the emitted binary. Every
/// caller was checked — direct and through the `set_*_const`/`set_*_from_value`
/// helpers — and only `Null`/`Bool`/`I32`/`I64`/`F64`/`String`/`BigInt` occur.
pub fn push_const(chunk: &mut Chunk, value: Value, line: u32) {
    match value {
        // Both push `Value::Null`; this is the same push through the spec
        // instruction instead of a pool index.
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::Bool(b) => chunk.emit_bool_const(b, line),
        Value::I32(n) => chunk.emit_i32_const(n, line),
        Value::I64(n) => chunk.emit_i64_const(n, line),
        Value::F32(n) => chunk.emit_f32_const(n, line),
        Value::F64(n) => chunk.emit_f64_const(n, line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        // `Value::Undefined` has no literal form; it is the global.
        Value::Undefined => {
            crate::primitives::globals::emit_read(chunk, "undefined", line);
        }
        // AST bigint literals always fit i64 — oversize ones are normalized to
        // `BigInt("…")` by the walker — so the ToBigInt64 wrap is lossless.
        Value::BigInt(v) => {
            chunk.emit_i64_const(v.to_i64_wrapping(), line);
            let idx = chunk.add_import("wasm:js-bigint", "fromI64");
            chunk.emit_call(idx, 1, line);
        }
        Value::V128(v) => {
            chunk.emit_op(Op::V128_CONST, line);
            for b in v {
                chunk.emit(b, line);
            }
        }
        other => panic!("push_const: no WASM-compliant encoding for {:?}", other),
    }
}

// ── Arithmetic ─────────────────────────────────────────────────────────────
//
// Everything below operates on RAW MILLISECONDS, never on a language's date
// object. That is what makes it centralisable: PHP's `{__time,__tz}` struct,
// Java's `Instant`, .NET's `DateTime` and Python's `datetime` all differ in
// shape and in mutability rules, so each adapter extracts a time value, calls
// in here, and stores the result back the way its own type requires.
//
// The split mirrors what PHP had already discovered independently:
//   * FIXED-unit adds (ms/s/min/hour/day/week) are pure scaling — no policy.
//   * CALENDAR adds (month/year) need `MonthOverflow`, because the answer to
//     `Jan 31 + 1 month` is a language decision, not a fact.

/// Millisecond spans. Eighteen literal occurrences of `86_400_000` across eight
/// adapter files (php, python, java, dotnet, dart, ruby, cobol, python-time)
/// is eight chances to type one digit wrong.
pub const MS_PER_SECOND: f64 = 1_000.0;
pub const MS_PER_MINUTE: f64 = 60_000.0;
pub const MS_PER_HOUR: f64 = 3_600_000.0;
pub const MS_PER_DAY: f64 = 86_400_000.0;
pub const MS_PER_WEEK: f64 = 604_800_000.0;
pub const DOTNET_DATETIME_MIN_UNIX_MS: f64 = -62_135_596_800_000.0;
pub const DOTNET_DATETIME_MAX_UNIX_MS: f64 = 253_402_300_799_999.0;
pub const DOTNET_TICKS_PER_MS: f64 = 10_000.0;
pub const DOTNET_TICKS_AT_UNIX_EPOCH: f64 = 621_355_968_000_000_000.0;

/// `ms + n * unit_ms` — every fixed-unit add and subtract. Subtraction is the
/// caller negating `n`, which is why there is no `emit_sub_scaled`.
///
/// Stack: `[ms, n]` → `[ms']`.
pub fn emit_add_scaled(chunk: &mut Chunk, unit_ms: f64, line: u32) {
    chunk.emit_f64_const(unit_ms, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
}

/// `ms + n` months, honouring the language's [`MonthOverflow`].
///
/// The one place `MonthOverflow` does real work. `2024-01-31 + 1 month` is
/// `2024-03-02` under `Overflow` (PHP, JavaScript — surplus days roll into the
/// next month, a consequence of `Date.prototype.setMonth`) and `2024-02-29`
/// under `Clamp` (Java, .NET, Python). Both are correct; they are different
/// languages' answers to the same question.
///
/// Years are months × 12, so callers add years through this too.
///
/// Stack: `[ms, n_months]` → `[ms']`.
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

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::datetime::{MonthOverflow, WeekdayBase};

    #[test]
    fn unit_spans_are_consistent() {
        assert_eq!(MS_PER_MINUTE, MS_PER_SECOND * 60.0);
        assert_eq!(MS_PER_HOUR, MS_PER_MINUTE * 60.0);
        assert_eq!(MS_PER_DAY, MS_PER_HOUR * 24.0);
        assert_eq!(MS_PER_WEEK, MS_PER_DAY * 7.0);
        assert_eq!(DOTNET_TICKS_PER_MS, 10_000.0);
    }

    /// The two month-overflow answers must actually differ, or the policy is
    /// decorative. PHP/JS overflow; Java/.NET/Python clamp.
    #[test]
    fn month_overflow_policies_disagree() {
        assert!(!month_add_clamps(DateTimePolicy::ECMA));
        assert!(month_add_clamps(DateTimePolicy::ISO));
        assert_eq!(DateTimePolicy::ECMA.month_overflow, MonthOverflow::Overflow);
        assert_eq!(DateTimePolicy::ISO.month_overflow, MonthOverflow::Clamp);
    }

    /// The ECMA preset must describe JavaScript, since it is the `Default` and
    /// everything compiles onto `ecma:date`.
    #[test]
    fn ecma_preset_matches_javascript() {
        let p = DateTimePolicy::default();
        assert_eq!(p.weekday_base, WeekdayBase::SundayZero);
        assert_eq!(p.month_indexing, MonthIndexing::ZeroBased);
        assert_eq!(p.epoch_precision, EpochPrecision::Millis);
    }
}
