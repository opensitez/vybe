use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Intl.DateTimeFormat` Timezone, Calendar & Date Formatting Options
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_intl_datetimeformat_basic_date_formatting() {
    let src = r#"
const date = new Date(Date.UTC(2026, 6, 22));
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC" });
console.log(formatter.format(date));
"#;
    assert_eq!(run_js(src), vec!["7/22/2026"]);
}

#[test]
fn test_js_intl_datetimeformat_full_date_and_time_style() {
    let src = r#"
const date = new Date(Date.UTC(2026, 6, 22, 12, 0, 0));
const formatter = new Intl.DateTimeFormat("en-US", { dateStyle: "full", timeZone: "UTC" });
console.log(formatter.format(date));
"#;
    assert_eq!(run_js(src), vec!["Wednesday, July 22, 2026"]);
}

#[test]
fn test_js_intl_datetimeformat_timezone_option_utc() {
    let src = r#"
const date = new Date(Date.UTC(2026, 0, 1, 0, 0, 0));
const fmtUTC = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", hour: "numeric", minute: "numeric" });
console.log(fmtUTC.format(date));
"#;
    assert_eq!(run_js(src), vec!["12:00 AM"]);
}

#[test]
fn test_js_intl_datetimeformat_format_to_parts() {
    let src = r#"
const date = new Date(Date.UTC(2026, 6, 22));
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", year: "numeric", month: "numeric", day: "numeric" });
const parts = formatter.formatToParts(date);
console.log(parts.map(p => `${p.type}:${p.value}`).join("|"));
"#;
    assert_eq!(
        run_js(src),
        vec!["month:7|literal:/|day:22|literal:/|year:2026"]
    );
}

#[test]
fn test_js_intl_datetimeformat_format_range() {
    let src = r#"
const d1 = new Date(Date.UTC(2026, 6, 20));
const d2 = new Date(Date.UTC(2026, 6, 22));
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", month: "short", day: "numeric" });
console.log(formatter.formatRange(d1, d2));
"#;
    assert_eq!(run_js(src), vec!["Jul 20 – 22"]);
}

#[test]
fn test_js_intl_datetimeformat_calendar_option_gregory() {
    let src = r#"
const formatter = new Intl.DateTimeFormat("en-US", { calendar: "gregory" });
console.log(formatter.resolvedOptions().calendar);
"#;
    assert_eq!(run_js(src), vec!["gregory"]);
}

#[test]
fn test_js_intl_datetimeformat_hour12_boolean_option() {
    let src = r#"
const date = new Date(Date.UTC(2026, 0, 1, 14, 30, 0));
const fmt12 = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", hour: "numeric", hour12: true });
const fmt24 = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", hour: "numeric", hour12: false });
console.log(fmt12.format(date) + "|" + fmt24.format(date));
"#;
    assert_eq!(run_js(src), vec!["2 PM|14"]);
}

#[test]
fn test_js_intl_datetimeformat_fractional_second_digits() {
    let src = r#"
const date = new Date(Date.UTC(2026, 0, 1, 0, 0, 0, 123));
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", second: "numeric", fractionalSecondDigits: 3 });
console.log(formatter.format(date));
"#;
    assert_eq!(run_js(src), vec!["0.123"]);
}

#[test]
fn test_js_intl_datetimeformat_invalid_timezone_throws_rangeerror() {
    let src = r#"
try {
    new Intl.DateTimeFormat("en-US", { timeZone: "Invalid/Timezone_Name" });
} catch (e) {
    console.log("Invalid TimeZone RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid TimeZone RangeError"]);
}

#[test]
fn test_js_intl_datetimeformat_supported_locales_of() {
    let src = r#"
const locales = Intl.DateTimeFormat.supportedLocalesOf(["en-US", "ja-JP"]);
console.log(locales.length);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_intl_datetimeformat_timestamp_number_input() {
    let src = r#"
const timestamp = 1784678400000; // Epoch timestamp
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", year: "numeric" });
console.log(formatter.format(timestamp));
"#;
    assert_eq!(run_js(src), vec!["2026"]);
}

#[test]
fn test_js_intl_datetimeformat_year_two_digit() {
    let src = r#"
const date = new Date(Date.UTC(2026, 6, 22));
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", year: "2-digit" });
console.log(formatter.format(date));
"#;
    assert_eq!(run_js(src), vec!["26"]);
}

#[test]
fn test_js_intl_datetimeformat_weekday_long() {
    let src = r#"
const date = new Date(Date.UTC(2026, 6, 22));
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", weekday: "long" });
console.log(formatter.format(date));
"#;
    assert_eq!(run_js(src), vec!["Wednesday"]);
}

#[test]
fn test_js_intl_datetimeformat_month_long() {
    let src = r#"
const date = new Date(Date.UTC(2026, 6, 22));
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", month: "long" });
console.log(formatter.format(date));
"#;
    assert_eq!(run_js(src), vec!["July"]);
}

#[test]
fn test_js_intl_datetimeformat_timezone_name_short() {
    let src = r#"
const date = new Date(Date.UTC(2026, 6, 22));
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", timeZoneName: "short" });
console.log(formatter.format(date).includes("UTC"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_datetimeformat_format_range_to_parts() {
    let src = r#"
const d1 = new Date(Date.UTC(2026, 6, 20));
const d2 = new Date(Date.UTC(2026, 6, 22));
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", day: "numeric" });
const parts = formatter.formatRangeToParts(d1, d2);
console.log(parts.some(p => p.type === "day"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_datetimeformat_invalid_date_input_throws_rangeerror() {
    let src = r#"
const invalidDate = new Date(NaN);
const formatter = new Intl.DateTimeFormat("en-US");
try {
    formatter.format(invalidDate);
} catch (e) {
    console.log("Invalid Date Format RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Date Format RangeError"]);
}

#[test]
fn test_js_intl_datetimeformat_resolved_options_keys() {
    let src = r#"
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC" });
const opts = formatter.resolvedOptions();
console.log(opts.locale + "|" + opts.timeZone);
"#;
    assert_eq!(run_js(src), vec!["en-US|UTC"]);
}

#[test]
fn test_js_intl_datetimeformat_dayperiod_short() {
    let src = r#"
const date = new Date(Date.UTC(2026, 0, 1, 8, 0, 0));
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", hour: "numeric", dayPeriod: "short" });
console.log(formatter.format(date).includes("morning") || formatter.format(date).includes("AM"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_datetimeformat_numeric_vs_2digit_day() {
    let src = r#"
const date = new Date(Date.UTC(2026, 6, 5));
const fmtNum = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", day: "numeric" });
const fmt2Digit = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", day: "2-digit" });
console.log(fmtNum.format(date) + "|" + fmt2Digit.format(date));
"#;
    assert_eq!(run_js(src), vec!["5|05"]);
}
