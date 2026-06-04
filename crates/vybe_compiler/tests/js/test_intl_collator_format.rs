/// Intl.Collator, Intl.NumberFormat, Intl.DateTimeFormat deep dive —
/// locale-sensitive sorting, number formatting options, date formatting parts.
use super::helpers::run_js;

// ── Intl.Collator ─────────────────────────────────────────────────────────────

#[test]
fn collator_sorts_locale_sensitive() {
    assert_eq!(
        run_js(
            r#"
const words = ["Zebra", "apple", "Banana"];
const sorted = words.sort(new Intl.Collator("en", { sensitivity: "base" }).compare);
// Case-insensitive sort: apple, Banana, Zebra
console.log(sorted[0].toLowerCase());
"#
        ),
        vec!["apple"]
    );
}

#[test]
fn collator_numeric_ordering() {
    assert_eq!(
        run_js(
            r#"
const files = ["file10", "file2", "file1"];
const sorted = files.sort(new Intl.Collator("en", { numeric: true }).compare);
console.log(sorted.join(","));
"#
        ),
        vec!["file1,file2,file10"]
    );
}

#[test]
fn collator_case_first() {
    assert_eq!(
        run_js(
            r#"
const col = new Intl.Collator("en", { caseFirst: "upper" });
const result = col.compare("A", "a");
console.log(typeof result === "number");
"#
        ),
        vec!["true"]
    );
}

#[test]
fn collator_compare_returns_number() {
    assert_eq!(
        run_js(
            r#"
const col = new Intl.Collator("en");
const r1 = col.compare("apple", "banana");
const r2 = col.compare("banana", "apple");
const r3 = col.compare("same", "same");
console.log(r1 < 0);
console.log(r2 > 0);
console.log(r3 === 0);
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn collator_resolved_options() {
    assert_eq!(
        run_js(
            r#"
const col = new Intl.Collator("en-US");
const opts = col.resolvedOptions();
console.log(typeof opts.locale);
console.log(typeof opts.sensitivity);
"#
        ),
        vec!["string", "string"]
    );
}

// ── Intl.NumberFormat ─────────────────────────────────────────────────────────

#[test]
fn number_format_currency() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const result = fmt.format(1234.56);
console.log(result.includes("1,234.56"));
console.log(result.includes("$"));
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
const fmt = new Intl.NumberFormat("en-US", { style: "percent" });
console.log(fmt.format(0.85));
"#
        ),
        vec!["85%"]
    );
}

#[test]
fn number_format_max_fraction_digits() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 });
console.log(fmt.format(3.14159));
"#
        ),
        vec!["3.14"]
    );
}

#[test]
fn number_format_min_integer_digits() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en-US", { minimumIntegerDigits: 4 });
console.log(fmt.format(42));
"#
        ),
        vec!["0,042"]
    );
}

#[test]
fn number_format_significant_digits() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en-US", {
    minimumSignificantDigits: 3,
    maximumSignificantDigits: 4
});
console.log(fmt.format(1.2345));
"#
        ),
        vec!["1.235"]
    );
}

#[test]
fn number_format_to_parts() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const parts = fmt.formatToParts(1234.5);
const types = parts.map(p => p.type);
console.log(types.includes("currency"));
console.log(types.includes("integer"));
console.log(types.includes("decimal"));
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn number_format_notation_compact() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en-US", { notation: "compact" });
const result = fmt.format(1000000);
// Should be "1M" or similar
console.log(result.includes("M") || result.includes("m") || result.length < 8);
"#
        ),
        vec!["true"]
    );
}

// ── Intl.DateTimeFormat ───────────────────────────────────────────────────────

#[test]
fn datetimeformat_basic_date() {
    assert_eq!(
        run_js(
            r#"
const date = new Date(2024, 0, 15); // Jan 15, 2024
const fmt = new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "numeric" });
const result = fmt.format(date);
console.log(result.includes("2024"));
console.log(result.includes("January") || result.includes("Jan"));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn datetimeformat_format_to_parts() {
    assert_eq!(
        run_js(
            r#"
const date = new Date(2024, 5, 10);
const fmt = new Intl.DateTimeFormat("en-US", { year: "numeric", month: "2-digit", day: "2-digit" });
const parts = fmt.formatToParts(date);
const types = parts.map(p => p.type);
console.log(types.includes("year"));
console.log(types.includes("month"));
console.log(types.includes("day"));
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn datetimeformat_weekday_option() {
    assert_eq!(
        run_js(
            r#"
// June 10 2024 is a Monday
const date = new Date(2024, 5, 10);
const fmt = new Intl.DateTimeFormat("en-US", { weekday: "long" });
const result = fmt.format(date);
console.log(result === "Monday");
"#
        ),
        vec!["true"]
    );
}

#[test]
fn datetimeformat_resolved_options() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.DateTimeFormat("en-US");
const opts = fmt.resolvedOptions();
console.log(typeof opts.locale);
console.log(typeof opts.timeZone);
"#
        ),
        vec!["string", "string"]
    );
}

// ── Intl.ListFormat ───────────────────────────────────────────────────────────

#[test]
fn list_format_conjunction() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.ListFormat("en-US", { type: "conjunction" });
const result = fmt.format(["a", "b", "c"]);
console.log(result.includes("and"));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn list_format_disjunction() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.ListFormat("en-US", { type: "disjunction" });
const result = fmt.format(["cats", "dogs"]);
console.log(result.includes("or"));
"#
        ),
        vec!["true"]
    );
}

// ── Intl.PluralRules ──────────────────────────────────────────────────────────

#[test]
fn plural_rules_cardinal_english() {
    assert_eq!(
        run_js(
            r#"
const pr = new Intl.PluralRules("en-US");
console.log(pr.select(1));
console.log(pr.select(2));
console.log(pr.select(0));
"#
        ),
        vec!["one", "other", "other"]
    );
}

#[test]
fn plural_rules_ordinal() {
    assert_eq!(
        run_js(
            r#"
const pr = new Intl.PluralRules("en-US", { type: "ordinal" });
console.log(pr.select(1));  // "one" → 1st
console.log(pr.select(2));  // "two" → 2nd
console.log(pr.select(3));  // "few" → 3rd
"#
        ),
        vec!["one", "two", "few"]
    );
}

// ── Intl.RelativeTimeFormat ───────────────────────────────────────────────────

#[test]
fn relative_time_format_days() {
    assert_eq!(
        run_js(
            r#"
const rtf = new Intl.RelativeTimeFormat("en-US", { numeric: "auto" });
console.log(rtf.format(-1, "day"));
console.log(rtf.format(1, "day"));
"#
        ),
        vec!["yesterday", "tomorrow"]
    );
}

#[test]
fn relative_time_format_numeric() {
    assert_eq!(
        run_js(
            r#"
const rtf = new Intl.RelativeTimeFormat("en-US", { numeric: "always" });
const result = rtf.format(-3, "day");
console.log(result.includes("3"));
console.log(result.includes("ago"));
"#
        ),
        vec!["true", "true"]
    );
}
