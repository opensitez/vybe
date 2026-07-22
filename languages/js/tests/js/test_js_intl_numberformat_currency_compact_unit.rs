use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Intl.NumberFormat` Currency, Percent, Compact & Unit Formatting
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_intl_numberformat_decimal_formatting_default_locale() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US");
console.log(formatter.format(1234567.89));
"#;
    assert_eq!(run_js(src), vec!["1,234,567.89"]);
}

#[test]
fn test_js_intl_numberformat_currency_usd() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
console.log(formatter.format(100.5));
"#;
    assert_eq!(run_js(src), vec!["$100.50"]);
}

#[test]
fn test_js_intl_numberformat_percent_style() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { style: "percent" });
console.log(formatter.format(0.75));
"#;
    assert_eq!(run_js(src), vec!["75%"]);
}

#[test]
fn test_js_intl_numberformat_compact_display_short() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { notation: "compact", compactDisplay: "short" });
console.log(`${formatter.format(1000)}:${formatter.format(1000000)}`);
"#;
    assert_eq!(run_js(src), vec!["1K:1M"]);
}

#[test]
fn test_js_intl_numberformat_unit_formatting_speed() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { style: "unit", unit: "kilometer-per-hour" });
console.log(formatter.format(120));
"#;
    assert_eq!(run_js(src), vec!["120 km/h"]);
}

#[test]
fn test_js_intl_numberformat_format_to_parts() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const parts = formatter.formatToParts(1000.5);
console.log(parts.map(p => `${p.type}:${p.value}`).join("|"));
"#;
    assert_eq!(
        run_js(src),
        vec!["currency:$|integer:1|group:,|integer:000|decimal:.|fraction:50"]
    );
}

#[test]
fn test_js_intl_numberformat_format_range() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
console.log(formatter.formatRange(3, 5));
"#;
    assert_eq!(run_js(src), vec!["$3.00 – $5.00"]);
}

#[test]
fn test_js_intl_numberformat_minimum_fraction_digits() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { minimumFractionDigits: 3 });
console.log(formatter.format(4.2));
"#;
    assert_eq!(run_js(src), vec!["4.200"]);
}

#[test]
fn test_js_intl_numberformat_maximum_fraction_digits() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { maximumFractionDigits: 1 });
console.log(formatter.format(4.285));
"#;
    assert_eq!(run_js(src), vec!["4.3"]);
}

#[test]
fn test_js_intl_numberformat_supported_locales_of() {
    let src = r#"
const supported = Intl.NumberFormat.supportedLocalesOf(["en-US", "fr-FR"]);
console.log(supported.includes("en-US"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_numberformat_resolved_options() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { style: "currency", currency: "EUR" });
const opts = formatter.resolvedOptions();
console.log(opts.locale + "|" + opts.style + "|" + opts.currency);
"#;
    assert_eq!(run_js(src), vec!["en-US|currency|EUR"]);
}

#[test]
fn test_js_intl_numberformat_missing_currency_throws_typeerror() {
    let src = r#"
try {
    new Intl.NumberFormat("en-US", { style: "currency" }); // style: "currency" REQUIRES currency option!
} catch (e) {
    console.log("Currency Option Missing TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Currency Option Missing TypeError"]);
}

#[test]
fn test_js_intl_numberformat_use_grouping_boolean() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { useGrouping: false });
console.log(formatter.format(1000000));
"#;
    assert_eq!(run_js(src), vec!["1000000"]);
}

#[test]
fn test_js_intl_numberformat_bigint_formatting() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US");
console.log(formatter.format(1000000000000000n));
"#;
    assert_eq!(run_js(src), vec!["1,000,000,000,000,000"]);
}

#[test]
fn test_js_intl_numberformat_sign_display_always() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { signDisplay: "always" });
console.log(`${formatter.format(5)}:${formatter.format(-5)}:${formatter.format(0)}`);
"#;
    assert_eq!(run_js(src), vec!["+5:-5:+0"]);
}

#[test]
fn test_js_intl_numberformat_unit_display_long() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { style: "unit", unit: "meter", unitDisplay: "long" });
console.log(formatter.format(5));
"#;
    assert_eq!(run_js(src), vec!["5 meters"]);
}

#[test]
fn test_js_intl_numberformat_format_range_to_parts() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const parts = formatter.formatRangeToParts(10, 20);
console.log(parts.some(p => p.type === "currency"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_numberformat_invalid_locale_throws_rangeerror() {
    let src = r#"
try {
    new Intl.NumberFormat("invalid_locale_tag!@#");
} catch (e) {
    console.log("Invalid Locale Tag RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Locale Tag RangeError"]);
}

#[test]
fn test_js_intl_numberformat_rounding_mode_ceil() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0, roundingMode: "ceil" });
console.log(formatter.format(1.1));
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_intl_numberformat_accounting_currency_sign() {
    let src = r#"
const formatter = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", currencySign: "accounting" });
console.log(formatter.format(-100));
"#;
    assert_eq!(run_js(src), vec!["($100.00)"]);
}
