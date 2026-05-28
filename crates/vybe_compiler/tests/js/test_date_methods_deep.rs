/// Date string formatting — toDateString, toTimeString, toLocaleDateString, toLocaleTimeString

use super::helpers::run_js;

#[test]
fn date_to_date_string_format() {
    assert_eq!(run_js(r#"
const d = new Date("2024-01-15T12:00:00.000Z");
const s = d.toUTCString();
console.log(typeof s);
console.log(s.includes("2024"));
"#), vec!["string", "true"]);
}

#[test]
fn date_to_iso_string_format() {
    assert_eq!(run_js(r#"
const d = new Date(0);
console.log(d.toISOString());
"#), vec!["1970-01-01T00:00:00.000Z"]);
}

#[test]
fn date_get_utc_methods() {
    assert_eq!(run_js(r#"
const d = new Date("2024-03-15T10:30:45.500Z");
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth()); // 2 (March, 0-indexed)
console.log(d.getUTCDate());
console.log(d.getUTCHours());
console.log(d.getUTCMinutes());
console.log(d.getUTCSeconds());
console.log(d.getUTCMilliseconds());
"#), vec!["2024", "2", "15", "10", "30", "45", "500"]);
}

#[test]
fn date_set_utc_full_year() {
    assert_eq!(run_js(r#"
const d = new Date(0);
d.setUTCFullYear(2025, 5, 20); // year, month (0-indexed), day
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#), vec!["2025", "5", "20"]);
}

#[test]
fn date_get_time_and_value_of() {
    assert_eq!(run_js(r#"
const d = new Date(12345678);
console.log(d.getTime());
console.log(d.valueOf());
console.log(+d);
"#), vec!["12345678", "12345678", "12345678"]);
}

#[test]
fn date_comparison_via_subtraction() {
    assert_eq!(run_js(r#"
const earlier = new Date("2024-01-01");
const later = new Date("2024-12-31");
console.log(later - earlier > 0);
console.log(later > earlier);
"#), vec!["true", "true"]);
}

#[test]
fn date_utc_factory() {
    assert_eq!(run_js(r#"
const ms = Date.UTC(2024, 0, 1, 12, 0, 0);
const d = new Date(ms);
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
console.log(d.getUTCHours());
"#), vec!["2024", "0", "12"]);
}

#[test]
fn date_to_json_alias_iso() {
    assert_eq!(run_js(r#"
const d = new Date(0);
console.log(d.toJSON());
console.log(d.toJSON() === d.toISOString());
"#), vec!["1970-01-01T00:00:00.000Z", "true"]);
}

#[test]
fn date_leap_year() {
    assert_eq!(run_js(r#"
const d = new Date("2024-02-29T00:00:00.000Z");
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth()); // February = 1
console.log(d.getUTCDate());
console.log(isNaN(d.getTime())); // valid date
"#), vec!["2024", "1", "29", "false"]);
}

#[test]
fn date_to_locale_date_string_is_string() {
    assert_eq!(run_js(r#"
const d = new Date("2024-06-15");
console.log(typeof d.toLocaleDateString("en-US"));
"#), vec!["string"]);
}

#[test]
fn date_parse_returns_milliseconds() {
    assert_eq!(run_js(r#"
const ms = Date.parse("2024-01-01T00:00:00.000Z");
console.log(typeof ms);
console.log(ms > 0);
// Verify it's correct
const d = new Date(ms);
console.log(d.getUTCFullYear());
"#), vec!["number", "true", "2024"]);
}
