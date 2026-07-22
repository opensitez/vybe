use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Intl.RelativeTimeFormat` Relative Time Formatting
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_intl_relativetimeformat_past_and_future_days() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en", { numeric: "always" });
console.log(rtf.format(-1, "day") + "|" + rtf.format(1, "day"));
"#;
    assert_eq!(run_js(src), vec!["1 day ago|in 1 day"]);
}

#[test]
fn test_js_intl_relativetimeformat_auto_numeric_yesterday_tomorrow() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en", { numeric: "auto" });
console.log(rtf.format(-1, "day") + "|" + rtf.format(1, "day") + "|" + rtf.format(0, "day"));
"#;
    assert_eq!(run_js(src), vec!["yesterday|tomorrow|today"]);
}

#[test]
fn test_js_intl_relativetimeformat_units_seconds_minutes_hours() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en", { numeric: "always" });
console.log(`${rtf.format(-30, "second")}:${rtf.format(5, "minute")}:${rtf.format(-2, "hour")}`);
"#;
    assert_eq!(run_js(src), vec!["30 seconds ago:in 5 minutes:2 hours ago"]);
}

#[test]
fn test_js_intl_relativetimeformat_units_weeks_months_years() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en", { numeric: "always" });
console.log(`${rtf.format(-1, "week")}:${rtf.format(3, "month")}:${rtf.format(-10, "year")}`);
"#;
    assert_eq!(run_js(src), vec!["1 week ago:in 3 months:10 years ago"]);
}

#[test]
fn test_js_intl_relativetimeformat_format_to_parts() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en", { numeric: "always" });
const parts = rtf.formatToParts(-1, "day");
console.log(parts.map(p => `${p.type}:${p.value}`).join("|"));
"#;
    assert_eq!(run_js(src), vec!["integer:1|literal: day ago"]);
}

#[test]
fn test_js_intl_relativetimeformat_style_short() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en", { style: "short" });
console.log(rtf.format(-3, "month"));
"#;
    assert_eq!(run_js(src), vec!["3 mo. ago"]);
}

#[test]
fn test_js_intl_relativetimeformat_style_narrow() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en", { style: "narrow" });
console.log(rtf.format(2, "year"));
"#;
    assert_eq!(run_js(src), vec!["in 2 yr."]);
}

#[test]
fn test_js_intl_relativetimeformat_resolved_options() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en-US", { numeric: "auto", style: "short" });
const opts = rtf.resolvedOptions();
console.log(opts.locale + "|" + opts.numeric + "|" + opts.style);
"#;
    assert_eq!(run_js(src), vec!["en-US|auto|short"]);
}

#[test]
fn test_js_intl_relativetimeformat_supported_locales_of() {
    let src = r#"
const supported = Intl.RelativeTimeFormat.supportedLocalesOf(["en-US", "fr-FR"]);
console.log(supported.includes("en-US"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_relativetimeformat_invalid_unit_throws_rangeerror() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en");
try {
    rtf.format(1, "invalid_unit");
} catch (e) {
    console.log("Invalid Unit RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Unit RangeError"]);
}

#[test]
fn test_js_intl_relativetimeformat_plural_unit_aliases() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en");
console.log(rtf.format(5, "days") === rtf.format(5, "day"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_relativetimeformat_number_coercion() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en");
console.log(rtf.format("-5", "day"));
"#;
    assert_eq!(run_js(src), vec!["5 days ago"]);
}

#[test]
fn test_js_intl_relativetimeformat_zero_value_formatting() {
    let src = r#"
const rtfAlways = new Intl.RelativeTimeFormat("en", { numeric: "always" });
console.log(rtfAlways.format(0, "second") + "|" + rtfAlways.format(-0, "second"));
"#;
    assert_eq!(run_js(src), vec!["in 0 seconds|0 seconds ago"]);
}

#[test]
fn test_js_intl_relativetimeformat_quarter_unit() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en", { numeric: "always" });
console.log(rtf.format(1, "quarter") + "|" + rtf.format(-1, "quarter"));
"#;
    assert_eq!(run_js(src), vec!["in 1 quarter|1 quarter ago"]);
}

#[test]
fn test_js_intl_relativetimeformat_auto_numeric_quarter_this_last_next() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en", { numeric: "auto" });
console.log(rtf.format(-1, "quarter") + "|" + rtf.format(1, "quarter") + "|" + rtf.format(0, "quarter"));
"#;
    assert_eq!(run_js(src), vec!["last quarter|next quarter|this quarter"]);
}

#[test]
fn test_js_intl_relativetimeformat_invalid_style_throws_rangeerror() {
    let src = r#"
try {
    new Intl.RelativeTimeFormat("en", { style: "invalid_style" });
} catch (e) {
    console.log("Invalid Style RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Style RangeError"]);
}

#[test]
fn test_js_intl_relativetimeformat_symbol_number_argument_throws_typeerror() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en");
try {
    rtf.format(Symbol("1"), "day");
} catch (e) {
    console.log("RelativeTimeFormat Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["RelativeTimeFormat Symbol TypeError"]);
}

#[test]
fn test_js_intl_relativetimeformat_format_to_parts_structure() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en", { numeric: "always" });
const parts = rtf.formatToParts(10, "minute");
console.log(parts.some(p => p.type === "unit" && p.value === " minutes"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_relativetimeformat_auto_numeric_now_seconds() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("en", { numeric: "auto" });
console.log(rtf.format(0, "second"));
"#;
    assert_eq!(run_js(src), vec!["now"]);
}

#[test]
fn test_js_intl_relativetimeformat_locale_canonicalization() {
    let src = r#"
const rtf = new Intl.RelativeTimeFormat("EN-us");
console.log(rtf.resolvedOptions().locale);
"#;
    assert_eq!(run_js(src), vec!["en-US"]);
}
