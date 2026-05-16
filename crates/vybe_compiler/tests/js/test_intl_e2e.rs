//! End-to-end Intl tests — verifies that JS code calling `new Intl.X()`
//! and `instance.method()` resolves through the full pipeline:
//!
//!   1. JS profile parses `new Intl.NumberFormat(...)` → namespace lookup
//!   2. Namespace registry (`namespaces/intl.rs`) → constructor host fn
//!   3. Constructor returns Object stamped with `__type=NumberFormat`
//!   4. `nf.format(123)` → `STRUCT_GET "format"` → TypeRegistry vtable
//!   5. TypeRegistry routes to `ecma:intl/numberformat:format` host fn
//!   6. Host fn returns ICU-formatted string via writeable
//!
//! These exercise the same Component Model + ESM integration paths a
//! real Node/Deno program would use, just with Vybe-side ICU.

use crate::helpers::run_js;

fn run_js_one(src: &str) -> String {
    run_js(src).join(" ")
}

#[test]
fn intl_number_format_default_decimal() {
    let out = run_js_one(r#"
        const nf = new Intl.NumberFormat();
        console.log(nf.format(1234.5));
    "#);
    assert_eq!(out, "1,234.5");
}

#[test]
fn intl_number_format_german_locale() {
    let out = run_js_one(r#"
        const nf = new Intl.NumberFormat("de-DE");
        console.log(nf.format(1234.5));
    "#);
    // German: "1.234,5"
    assert_eq!(out, "1.234,5");
}

#[test]
fn intl_number_format_currency() {
    let out = run_js_one(r#"
        const nf = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
        console.log(nf.format(99.5));
    "#);
    // en-US currency: "$99.50"
    assert!(out.contains("99"), "Expected currency value, got {:?}", out);
    assert!(out.contains("$"), "Expected $ symbol, got {:?}", out);
}

#[test]
fn intl_collator_compare() {
    let out = run_js_one(r#"
        const c = new Intl.Collator();
        console.log(c.compare("a", "b"));
    "#);
    // Should print -1 (a < b)
    assert_eq!(out, "-1");
}

#[test]
fn intl_plural_rules_select() {
    let out = run_js_one(r#"
        const pr = new Intl.PluralRules("en");
        console.log(pr.select(1), pr.select(2));
    "#);
    assert_eq!(out, "one other");
}

#[test]
fn intl_plural_rules_russian() {
    let out = run_js_one(r#"
        const pr = new Intl.PluralRules("ru");
        console.log(pr.select(2));
    "#);
    // Russian: 2 → "few"
    assert_eq!(out, "few");
}

#[test]
fn intl_list_format_french() {
    let out = run_js_one(r#"
        const lf = new Intl.ListFormat("fr");
        console.log(lf.format(["a", "b", "c"]));
    "#);
    // French: "a, b et c"
    assert!(out.contains("et"), "Expected French 'et', got {:?}", out);
}

#[test]
fn intl_relative_time_format() {
    let out = run_js_one(r#"
        const rtf = new Intl.RelativeTimeFormat("en");
        console.log(rtf.format(-3, "day"));
    "#);
    assert_eq!(out, "3 days ago");
}

#[test]
fn intl_display_names_french_for_en() {
    let out = run_js_one(r#"
        const dn = new Intl.DisplayNames("en", { type: "language" });
        console.log(dn.of("fr"));
    "#);
    assert_eq!(out, "French");
}

#[test]
fn intl_locale_parse_complex_tag() {
    let out = run_js_one(r#"
        const loc = new Intl.Locale("zh-Hans-CN");
        console.log(loc.language, loc.script, loc.region);
    "#);
    assert_eq!(out, "zh Hans CN");
}

#[test]
fn intl_segmenter_grapheme() {
    let out = run_js_one(r#"
        const seg = new Intl.Segmenter();
        const segments = seg.segment("hi");
        console.log(segments.length);
    "#);
    assert_eq!(out, "2");
}

#[test]
fn intl_get_canonical_locales_static() {
    let out = run_js_one(r#"
        const tags = Intl.getCanonicalLocales(["EN-us", "FR-fr"]);
        console.log(tags[0], tags[1]);
    "#);
    assert_eq!(out, "en-US fr-FR");
}

#[test]
fn intl_supported_values_of_static() {
    let out = run_js_one(r#"
        const currencies = Intl.supportedValuesOf("currency");
        console.log(currencies.includes("USD"));
    "#);
    assert_eq!(out, "true");
}

#[test]
fn intl_duration_format_arabic_dual() {
    let out = run_js_one(r#"
        const df = new Intl.DurationFormat("ar", { style: "long" });
        console.log(df.format({ hours: 2 }));
    "#);
    // Arabic dual form for 2 hours: "ساعتان"
    assert!(out.contains("ساعتان"),
        "Expected Arabic dual 'ساعتان' for 2 hours, got {:?}", out);
}

#[test]
fn intl_date_time_format_german() {
    let out = run_js_one(r#"
        const dtf = new Intl.DateTimeFormat("de-DE");
        console.log(dtf.format(1705276800000));
    "#);
    // German: contains 2024 + dot separator
    assert!(out.contains("2024") && out.contains("."),
        "Expected German date format, got {:?}", out);
}
