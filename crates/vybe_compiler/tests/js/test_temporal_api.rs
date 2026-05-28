/// Temporal API (Stage 3) — Temporal.PlainDate, PlainTime, PlainDateTime,
/// ZonedDateTime, Duration, Now, arithmetic, comparison, formatting.

use super::helpers::run_js;

// ── Temporal.PlainDate ────────────────────────────────────────────────────────

#[test]
fn plain_date_construction() {
    assert_eq!(run_js(r#"
const d = new Temporal.PlainDate(2024, 1, 15);
console.log(d.year);
console.log(d.month);
console.log(d.day);
"#), vec!["2024", "1", "15"]);
}

#[test]
fn plain_date_from_string() {
    assert_eq!(run_js(r#"
const d = Temporal.PlainDate.from("2024-03-21");
console.log(d.year);
console.log(d.month);
console.log(d.day);
"#), vec!["2024", "3", "21"]);
}

#[test]
fn plain_date_tostring() {
    assert_eq!(run_js(r#"
const d = new Temporal.PlainDate(2024, 3, 5);
console.log(d.toString());
"#), vec!["2024-03-05"]);
}

#[test]
fn plain_date_add_duration() {
    assert_eq!(run_js(r#"
const d = new Temporal.PlainDate(2024, 1, 15);
const d2 = d.add({ days: 10 });
console.log(d2.day);
console.log(d2.month);
"#), vec!["25", "1"]);
}

#[test]
fn plain_date_add_months() {
    assert_eq!(run_js(r#"
const d = new Temporal.PlainDate(2024, 1, 31);
const d2 = d.add({ months: 1 });
console.log(d2.month);
"#), vec!["2"]);
}

#[test]
fn plain_date_subtract() {
    assert_eq!(run_js(r#"
const d = new Temporal.PlainDate(2024, 3, 10);
const d2 = d.subtract({ days: 5 });
console.log(d2.day);
"#), vec!["5"]);
}

#[test]
fn plain_date_compare() {
    assert_eq!(run_js(r#"
const a = new Temporal.PlainDate(2024, 1, 1);
const b = new Temporal.PlainDate(2024, 6, 1);
console.log(Temporal.PlainDate.compare(a, b));
console.log(Temporal.PlainDate.compare(b, a));
console.log(Temporal.PlainDate.compare(a, a));
"#), vec!["-1", "1", "0"]);
}

#[test]
fn plain_date_until_duration() {
    assert_eq!(run_js(r#"
const start = new Temporal.PlainDate(2024, 1, 1);
const end = new Temporal.PlainDate(2024, 1, 11);
const dur = start.until(end);
console.log(dur.days);
"#), vec!["10"]);
}

#[test]
fn plain_date_since() {
    assert_eq!(run_js(r#"
const start = new Temporal.PlainDate(2024, 1, 1);
const end = new Temporal.PlainDate(2024, 1, 11);
const dur = end.since(start);
console.log(dur.days);
"#), vec!["10"]);
}

#[test]
fn plain_date_day_of_week() {
    assert_eq!(run_js(r#"
const d = Temporal.PlainDate.from("2024-01-01"); // Monday
console.log(d.dayOfWeek);
"#), vec!["1"]);
}

#[test]
fn plain_date_with_modification() {
    assert_eq!(run_js(r#"
const d = new Temporal.PlainDate(2024, 1, 15);
const d2 = d.with({ day: 1 });
console.log(d2.day);
console.log(d2.month);
console.log(d2.year);
"#), vec!["1", "1", "2024"]);
}

// ── Temporal.PlainTime ────────────────────────────────────────────────────────

#[test]
fn plain_time_construction() {
    assert_eq!(run_js(r#"
const t = new Temporal.PlainTime(10, 30, 45);
console.log(t.hour);
console.log(t.minute);
console.log(t.second);
"#), vec!["10", "30", "45"]);
}

#[test]
fn plain_time_from_string() {
    assert_eq!(run_js(r#"
const t = Temporal.PlainTime.from("14:30:00");
console.log(t.hour);
console.log(t.minute);
"#), vec!["14", "30"]);
}

#[test]
fn plain_time_tostring() {
    assert_eq!(run_js(r#"
const t = new Temporal.PlainTime(9, 5, 3);
console.log(t.toString());
"#), vec!["09:05:03"]);
}

#[test]
fn plain_time_add() {
    assert_eq!(run_js(r#"
const t = new Temporal.PlainTime(10, 30, 0);
const t2 = t.add({ hours: 2, minutes: 15 });
console.log(t2.hour);
console.log(t2.minute);
"#), vec!["12", "45"]);
}

// ── Temporal.PlainDateTime ────────────────────────────────────────────────────

#[test]
fn plain_datetime_construction() {
    assert_eq!(run_js(r#"
const dt = new Temporal.PlainDateTime(2024, 3, 15, 10, 30, 0);
console.log(dt.year);
console.log(dt.month);
console.log(dt.day);
console.log(dt.hour);
console.log(dt.minute);
"#), vec!["2024", "3", "15", "10", "30"]);
}

#[test]
fn plain_datetime_from_string() {
    assert_eq!(run_js(r#"
const dt = Temporal.PlainDateTime.from("2024-03-15T10:30:00");
console.log(dt.year);
console.log(dt.hour);
"#), vec!["2024", "10"]);
}

#[test]
fn plain_datetime_tostring() {
    assert_eq!(run_js(r#"
const dt = new Temporal.PlainDateTime(2024, 3, 5, 9, 5, 3);
console.log(dt.toString());
"#), vec!["2024-03-05T09:05:03"]);
}

// ── Temporal.Duration ─────────────────────────────────────────────────────────

#[test]
fn duration_construction() {
    assert_eq!(run_js(r#"
const dur = new Temporal.Duration(1, 2, 0, 3, 4, 5, 6);
console.log(dur.years);
console.log(dur.months);
console.log(dur.days);
console.log(dur.hours);
"#), vec!["1", "2", "3", "4"]);
}

#[test]
fn duration_from_object() {
    assert_eq!(run_js(r#"
const dur = Temporal.Duration.from({ days: 7, hours: 12 });
console.log(dur.days);
console.log(dur.hours);
"#), vec!["7", "12"]);
}

#[test]
fn duration_negate() {
    assert_eq!(run_js(r#"
const dur = Temporal.Duration.from({ days: 5 });
const neg = dur.negated();
console.log(neg.days);
"#), vec!["-5"]);
}

// ── Temporal.Now ─────────────────────────────────────────────────────────────

#[test]
fn temporal_now_plaindate_utc() {
    assert_eq!(run_js(r#"
const d = Temporal.Now.plainDateISO();
console.log(typeof d.year === "number");
console.log(d.year >= 2024);
"#), vec!["true", "true"]);
}

#[test]
fn temporal_now_instant_has_epoch_seconds() {
    assert_eq!(run_js(r#"
const now = Temporal.Now.instant();
console.log(typeof now.epochSeconds === "number");
console.log(now.epochSeconds > 1700000000);
"#), vec!["true", "true"]);
}

// ── Temporal.Instant ──────────────────────────────────────────────────────────

#[test]
fn instant_from_epoch_milliseconds() {
    assert_eq!(run_js(r#"
const inst = Temporal.Instant.fromEpochMilliseconds(0);
console.log(inst.epochSeconds);
"#), vec!["0"]);
}

#[test]
fn instant_compare() {
    assert_eq!(run_js(r#"
const a = Temporal.Instant.fromEpochMilliseconds(1000);
const b = Temporal.Instant.fromEpochMilliseconds(2000);
console.log(Temporal.Instant.compare(a, b));
"#), vec!["-1"]);
}
