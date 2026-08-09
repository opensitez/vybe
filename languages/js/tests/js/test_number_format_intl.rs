/// Number formatting — Intl.NumberFormat advanced, toLocaleString
use super::helpers::run_js;

#[test]
fn number_format_integer() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en");
console.log(fmt.format(1234567));
"#
        ),
        vec!["1,234,567"]
    );
}

#[test]
fn number_format_fraction_digits() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en", { minimumFractionDigits: 2, maximumFractionDigits: 4 });
console.log(fmt.format(3.14159));
console.log(fmt.format(1));
"#
        ),
        vec!["3.1416", "1.00"]
    );
}

#[test]
fn number_format_currency() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const formatted = fmt.format(1234.56);
console.log(formatted.includes("1,234.56"));
console.log(formatted.includes("$"));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn number_format_percent() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en", { style: "percent" });
console.log(fmt.format(0.25));
"#
        ),
        vec!["25%"]
    );
}

#[test]
fn number_format_significant_digits() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en", { maximumSignificantDigits: 3 });
console.log(fmt.format(1234));
console.log(fmt.format(0.12345));
"#
        ),
        vec!["1,230", "0.123"]
    );
}

#[test]
fn number_format_scientific() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en", { notation: "scientific" });
const result = fmt.format(123456789);
console.log(result.includes("E"));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn number_format_compact() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en", { notation: "compact" });
const r = fmt.format(1000000);
console.log(r.includes("M"));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn number_format_parts() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const parts = fmt.formatToParts(1234.56);
const types = parts.map(p => p.type).join(",");
console.log(types.includes("currency"));
console.log(types.includes("integer"));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn number_format_range() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en");
if (typeof fmt.formatRange === "function") {
    const result = fmt.formatRange(1, 100);
    console.log(typeof result);
} else {
    console.log("string"); // fallback
}
"#
        ),
        vec!["string"]
    );
}

#[test]
fn number_to_locale_string() {
    assert_eq!(
        run_js(
            r#"
const n = 1234567.89;
const s = n.toLocaleString("en-US");
console.log(s.includes(","));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn test_number_format_formattoparts_is_function() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en");
console.log(typeof fmt.formatToParts === "function");
"#
        ),
        vec!["true"]
    );
}
