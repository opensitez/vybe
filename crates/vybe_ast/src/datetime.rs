//! Calendar and time-zone POLICY vocabulary.
//!
//! The arithmetic of dates is identical everywhere — proleptic Gregorian, one
//! set of leap-year rules, one definition of "day of year". What differs
//! between languages is a small number of *decisions* taken at the edges, and
//! those decisions are known to the WALKER (they are properties of the source
//! language) but unrecoverable by the compiler from a method's spelling.
//!
//! That is the same shape as [`crate::class_normalize::types::AugmentationPolicy`]:
//! a language declares its mechanism ONCE as a constant, and the shared emitter
//! takes it as a parameter. Without it, `emit_add_months` would have to ask
//! "which language am I compiling", which is exactly the check that is banned
//! in shared code — and in practice it produced four separate implementations
//! of the same calendar maths (PHP, Python, Java, .NET), each with its own
//! `emit_days_in_month`.
//!
//! NOT represented here: anything the existing AST already carries. Date
//! construction is `ExprKind::New`, date literals are strings. A `DateTime`
//! node would force every walker to special-case what ordinary nodes already
//! express.

/// What happens when adding months lands on a day the target month does not
/// have — `Jan 31 + 1 month`.
///
/// This is a real divergence, not a subtlety: PHP and JavaScript OVERFLOW
/// (`2024-01-31 +1 month` → `2024-03-02`, because February has 29 days in 2024
/// and the extra 2 spill over), while Java `LocalDate.plusMonths`, .NET
/// `AddMonths` and Python `dateutil.relativedelta` CLAMP to `2024-02-29`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MonthOverflow {
    /// Pin to the last valid day of the target month. Java, .NET, Python.
    Clamp,
    /// Let surplus days roll into the following month. PHP, JavaScript —
    /// a consequence of `Date.prototype.setMonth` semantics.
    Overflow }

/// How a local wall-clock time that does NOT EXIST is resolved — the hour
/// skipped when daylight saving begins.
///
/// A language must answer this to convert local time to an instant at all, and
/// they disagree: Java's `ZonedDateTime` shifts forward by the gap length,
/// Python's `fold` machinery and PHP's parser make their own choices, and some
/// APIs reject outright.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DstGap {
    /// Move forward by the gap (Java `ZonedDateTime.of`).
    ShiftForward,
    /// Move backward by the gap.
    ShiftBackward,
    /// Refuse — the local time is not a real instant.
    Reject }

/// How a local wall-clock time that occurs TWICE is resolved — the hour
/// repeated when daylight saving ends.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DstAmbiguous {
    /// The first (still-daylight) occurrence. Java's default, Python `fold=0`.
    Earlier,
    /// The second (standard-time) occurrence. Python `fold=1`.
    Later,
    /// Refuse rather than silently pick.
    Reject }

/// Which week-numbering rule `weekOfYear` follows.
///
/// Not cosmetic: ISO-8601 weeks start Monday and week 1 is the one containing
/// the first Thursday, so 1 January is frequently in week 52 or 53 OF THE
/// PREVIOUS YEAR — which is why an ISO week number must be paired with an ISO
/// week-year rather than the calendar year.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WeekNumbering {
    /// ISO-8601: Monday first, ≥4 days in the first week.
    Iso,
    /// Sunday first, week 1 contains 1 January. US convention.
    Us,
    /// Week 1 starts 1 January regardless of weekday; no week-year concept.
    DayOfYearBased }

/// The unit a language's epoch timestamps are counted in. Java `Instant` is
/// nanosecond-precision, JavaScript and PHP `DateTime` are milliseconds, Unix
/// `time_t` and PHP's `time()` are whole seconds.
///
/// Shared arithmetic works in MILLISECONDS (the ECMA-262 time value); this
/// says what the language's own surface expects, so conversion happens once at
/// the boundary instead of being rediscovered per call site.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EpochPrecision {
    Seconds,
    Millis,
    Micros,
    Nanos }

/// Which weekday a language numbers first when it exposes a numeric weekday.
///
/// JavaScript `getDay` and PHP `date('w')` are Sunday=0; ISO-8601, PHP
/// `date('N')` and Java `DayOfWeek` are Monday=1. Python `weekday()` is
/// Monday=0 — a third convention, which is why this is an enum and not a bool.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WeekdayBase {
    /// Sunday = 0 … Saturday = 6. JavaScript, PHP `w`.
    SundayZero,
    /// Monday = 1 … Sunday = 7. ISO-8601, Java, PHP `N`.
    MondayOne,
    /// Monday = 0 … Sunday = 6. Python `date.weekday()`.
    MondayZero }

/// Which integer a language uses for January.
///
/// JavaScript and everything built on `ecma:date` are 0-based (`getMonth()`
/// returns 0 for January); PHP's `date('n')`, Python, Java and .NET are
/// 1-based. This is the only genuine PARAMETER among the differences between
/// the four existing `days_in_month` implementations — the rest (slots vs
/// stack, i32 vs f64 return, date-object vs year input) are calling
/// conventions a wrapper absorbs, not semantics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MonthIndexing {
    /// January = 0. JavaScript, `ecma:date`.
    ZeroBased,
    /// January = 1. PHP, Python, Java, .NET, ISO-8601.
    OneBased }

/// A language's calendar decisions, declared ONCE by its walker/normalizer and
/// threaded into the shared emitters — the direct analogue of
/// `AugmentationPolicy`.
///
/// Every field is a decision some real language makes differently. Anything
/// where all languages agree (leap-year rule, days in each month, the
/// proleptic Gregorian calendar itself) is deliberately absent: it belongs in
/// the shared implementation, not in a per-language knob.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DateTimePolicy {
    pub month_overflow: MonthOverflow,
    pub dst_gap: DstGap,
    pub dst_ambiguous: DstAmbiguous,
    pub week_numbering: WeekNumbering,
    pub epoch_precision: EpochPrecision,
    pub weekday_base: WeekdayBase,
    pub month_indexing: MonthIndexing }

impl DateTimePolicy {
    /// ECMA-262 `Date` semantics — the baseline, because the runtime's time
    /// values ARE ECMA time values. A language that has not declared its own
    /// policy behaves like JavaScript, which is the least surprising default
    /// given everything compiles onto `ecma:date`.
    pub const ECMA: Self = Self {
        month_overflow: MonthOverflow::Overflow,
        dst_gap: DstGap::ShiftForward,
        dst_ambiguous: DstAmbiguous::Earlier,
        week_numbering: WeekNumbering::Us,
        epoch_precision: EpochPrecision::Millis,
        weekday_base: WeekdayBase::SundayZero,
        month_indexing: MonthIndexing::ZeroBased };

    /// ISO-8601 / `java.time` semantics: clamping month arithmetic and ISO
    /// week numbering. Also the closest fit for .NET and Python.
    pub const ISO: Self = Self {
        month_overflow: MonthOverflow::Clamp,
        dst_gap: DstGap::ShiftForward,
        dst_ambiguous: DstAmbiguous::Earlier,
        week_numbering: WeekNumbering::Iso,
        epoch_precision: EpochPrecision::Millis,
        weekday_base: WeekdayBase::MondayOne,
        month_indexing: MonthIndexing::OneBased };
}

impl Default for DateTimePolicy {
    fn default() -> Self {
        Self::ECMA
    }
}

use crate::{BinOp, ExprKind, Expression};

/// The proleptic-Gregorian leap rule EVALUATED, for callers that have the year
/// as a plain number rather than as bytecode or AST.
///
/// This is the third rendering of the one rule (see [`leap_year_expr`] for why
/// several exist). It serves the layer the other two cannot reach: a walker
/// folding a *literal* date at parse time, and host-side runtime code doing
/// calendar arithmetic in Rust. Both previously kept private copies.
pub fn is_leap_year(year: i64) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

/// Days in `month` of `year`, months numbered 1..=12.
///
/// One-based because that is how a calendar is written and how every source
/// language spells it; the zero-based form belongs to the JS `Date` API and is
/// handled by [`MonthIndexing`] on the bytecode side. Out-of-range months
/// return 30 — matching what the language copies of this did, so folding an
/// already-invalid date stays a validation decision for the caller rather than
/// a panic here.
pub fn days_in_month(year: i64, month: i64) -> i64 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if is_leap_year(year) => 29,
        2 => 28,
        _ => 30 }
}

/// The proleptic-Gregorian leap rule as an AST EXPRESSION:
/// `y % 4 == 0 && (y % 100 != 0 || y % 400 == 0)`.
///
/// The same rule as `primitives::datetime::emit_is_leap_year`, in the other
/// rendering. Two renderings exist because the two consumers live at different
/// layers, not because the rule differs: an *adapter* emits bytecode into a
/// `Chunk`, while a *walker* rewrites source into AST before any chunk exists.
/// Pascal's `IsLeapYear` is walker-lowered, so a bytecode emitter is unreachable
/// from it — and the walker form has an advantage the emitter cannot offer:
/// when the year is a literal, ordinary constant folding sees straight through
/// it.
///
/// Keep the two in step. If the rule ever changes (it will not — it is fixed by
/// the Gregorian calendar), it changes in both.
pub fn leap_year_expr(year: Expression) -> Expression {
    let bin = |op: BinOp, left: Expression, right: Expression| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right) })
    };
    let divisible_by = |divisor: i64| {
        bin(
            BinOp::Eq,
            bin(BinOp::Mod, year.clone(), Expression::int(divisor)),
            Expression::int(0),
        )
    };
    bin(
        BinOp::And,
        divisible_by(4),
        bin(
            BinOp::Or,
            bin(BinOp::Eq, divisible_by(100), Expression::bool(false)),
            divisible_by(400),
        ),
    )
}
