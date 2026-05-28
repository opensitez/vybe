/// Date object — construction, UTC, arithmetic, mutation, ISO string,
/// comparison, rollover, invalid date, Date.now, day-of-week.

use super::helpers::run_js;

#[test]
fn date_from_epoch_zero() {
    assert_eq!(run_js(r#"
const d = new Date(0);
console.log(d.getTime());
console.log(d.getUTCFullYear());
"#), vec!["0", "1970"]);
}

#[test]
fn date_from_iso_string() {
    assert_eq!(run_js(r#"
const d = new Date("2024-06-15T00:00:00.000Z");
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#), vec!["2024", "5", "15"]);
}

#[test]
fn date_from_components() {
    assert_eq!(run_js(r#"
const d = new Date(2024, 0, 15); // January (0-indexed)
console.log(d.getFullYear());
console.log(d.getMonth());
console.log(d.getDate());
"#), vec!["2024", "0", "15"]);
}

#[test]
fn invalid_date_nan_gettime() {
    assert_eq!(run_js(r#"
const d = new Date("not-a-date");
console.log(isNaN(d.getTime()));
"#), vec!["true"]);
}

#[test]
fn date_now_is_number() {
    assert_eq!(run_js(r#"
console.log(typeof Date.now());
console.log(Date.now() > 0);
"#), vec!["number", "true"]);
}

#[test]
fn date_arithmetic_add_days() {
    assert_eq!(run_js(r#"
const d = new Date("2024-01-01T00:00:00.000Z");
const next = new Date(d.getTime() + 7 * 24 * 60 * 60 * 1000);
console.log(next.getUTCDate());
"#), vec!["8"]);
}

#[test]
fn date_diff_in_days() {
    assert_eq!(run_js(r#"
const a = new Date("2024-01-01");
const b = new Date("2024-01-31");
const days = (b - a) / (1000 * 60 * 60 * 24);
console.log(days);
"#), vec!["30"]);
}

#[test]
fn date_comparison_operators() {
    assert_eq!(run_js(r#"
const a = new Date("2024-01-01");
const b = new Date("2024-06-01");
console.log(a < b);
console.log(b > a);
"#), vec!["true", "true"]);
}

#[test]
fn date_set_methods() {
    assert_eq!(run_js(r#"
const d = new Date("2024-01-15T00:00:00.000Z");
d.setUTCFullYear(2025);
d.setUTCMonth(5);
d.setUTCDate(20);
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#), vec!["2025", "5", "20"]);
}

#[test]
fn date_to_iso_string_epoch() {
    assert_eq!(run_js(r#"
const d = new Date(0);
console.log(d.toISOString());
"#), vec!["1970-01-01T00:00:00.000Z"]);
}

#[test]
fn date_tojson_matches_iso() {
    assert_eq!(run_js(r#"
const d = new Date(0);
console.log(d.toJSON() === d.toISOString());
"#), vec!["true"]);
}

#[test]
fn date_utc_static_factory() {
    assert_eq!(run_js(r#"
const ms = Date.UTC(2024, 0, 1);
const d = new Date(ms);
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
"#), vec!["2024", "0"]);
}

#[test]
fn date_parse_iso() {
    assert_eq!(run_js(r#"
const ms = Date.parse("2024-06-15T00:00:00.000Z");
console.log(typeof ms);
console.log(ms > 0);
"#), vec!["number", "true"]);
}

#[test]
fn date_month_overflow_rolls_year() {
    assert_eq!(run_js(r#"
const d = new Date(2024, 12, 1); // month 12 = January next year
console.log(d.getFullYear());
console.log(d.getMonth());
"#), vec!["2025", "0"]);
}

#[test]
fn date_day_overflow_rolls_month() {
    assert_eq!(run_js(r#"
const d = new Date(2024, 0, 32); // Jan 32 = Feb 1
console.log(d.getMonth());
console.log(d.getDate());
"#), vec!["1", "1"]);
}

#[test]
fn date_getday_is_zero_to_six() {
    assert_eq!(run_js(r#"
// 2024-01-07 is a Sunday (0)
const d = new Date("2024-01-07T12:00:00.000Z");
console.log(d.getUTCDay());
"#), vec!["0"]);
}

#[test]
fn date_set_hours_minutes_seconds() {
    assert_eq!(run_js(r#"
const d = new Date("2024-01-01T00:00:00.000Z");
d.setUTCHours(14, 30, 45);
console.log(d.getUTCHours());
console.log(d.getUTCMinutes());
console.log(d.getUTCSeconds());
"#), vec!["14", "30", "45"]);
}

#[test]
fn date_value_of_returns_timestamp() {
    assert_eq!(run_js(r#"
const d = new Date(12345);
console.log(d.valueOf());
console.log(+d);
"#), vec!["12345", "12345"]);
}
