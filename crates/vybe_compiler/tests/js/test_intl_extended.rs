use super::helpers::run_js;

// ── Intl.NumberFormat ─────────────────────────────────────
#[test]
fn intl_numberformat_basic() {
    assert_eq!(run_js(r#"
const fmt = new Intl.NumberFormat("en-US");
console.log(typeof fmt.format(1234));
"#), vec!["string"]);
}

#[test]
fn intl_numberformat_style_currency() {
    assert_eq!(run_js(r#"
const fmt = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const result = fmt.format(1234.5);
console.log(result.includes("1,234"));
"#), vec!["true"]);
}

#[test]
fn intl_numberformat_style_percent() {
    assert_eq!(run_js(r#"
const fmt = new Intl.NumberFormat("en-US", { style: "percent" });
const result = fmt.format(0.5);
console.log(result.includes("50"));
"#), vec!["true"]);
}

#[test]
fn intl_numberformat_maximumfractiondigits() {
    assert_eq!(run_js(r#"
const fmt = new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 });
const result = fmt.format(3.14159);
console.log(result.includes("3.14"));
"#), vec!["true"]);
}

#[test]
fn intl_numberformat_formattoparts() {
    assert_eq!(run_js(r#"
const fmt = new Intl.NumberFormat("en-US");
const parts = fmt.formatToParts(1234);
console.log(parts.some(p => p.type === "integer"));
"#), vec!["true"]);
}

// ── Intl.DateTimeFormat ───────────────────────────────────
#[test]
fn intl_datetimeformat_basic() {
    assert_eq!(run_js(r#"
const fmt = new Intl.DateTimeFormat("en-US");
const result = fmt.format(new Date(2024, 0, 15));
console.log(typeof result);
"#), vec!["string"]);
}

#[test]
fn intl_datetimeformat_year_month_day() {
    assert_eq!(run_js(r#"
const fmt = new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "numeric" });
const result = fmt.format(new Date(2024, 0, 15));
console.log(result.includes("2024"));
console.log(result.includes("January"));
"#), vec!["true", "true"]);
}

#[test]
fn intl_datetimeformat_formattoparts() {
    assert_eq!(run_js(r#"
const fmt = new Intl.DateTimeFormat("en-US", { year: "numeric" });
const parts = fmt.formatToParts(new Date(2024, 0, 1));
const year = parts.find(p => p.type === "year");
console.log(year.value);
"#), vec!["2024"]);
}

// ── Intl.Collator ─────────────────────────────────────────
#[test]
fn intl_collator_basic_compare() {
    assert_eq!(run_js(r#"
const coll = new Intl.Collator("en-US");
const words = ["banana", "apple", "cherry"];
words.sort(coll.compare);
console.log(words[0]);
"#), vec!["apple"]);
}

#[test]
fn intl_collator_case_sensitivity() {
    assert_eq!(run_js(r#"
const coll = new Intl.Collator("en-US", { sensitivity: "base" });
console.log(coll.compare("a", "A") === 0);
"#), vec!["true"]);
}

#[test]
fn intl_collator_compare_returns_number() {
    assert_eq!(run_js(r#"
const coll = new Intl.Collator("en-US");
const result = coll.compare("a", "b");
console.log(result < 0);
"#), vec!["true"]);
}

// ── Intl.PluralRules ──────────────────────────────────────
#[test]
fn intl_pluralrules_one_other() {
    assert_eq!(run_js(r#"
const pr = new Intl.PluralRules("en-US");
console.log(pr.select(1));
console.log(pr.select(2));
"#), vec!["one", "other"]);
}

#[test]
fn intl_pluralrules_ordinal() {
    assert_eq!(run_js(r#"
const pr = new Intl.PluralRules("en-US", { type: "ordinal" });
console.log(pr.select(1));
console.log(pr.select(2));
"#), vec!["one", "two"]);
}

// ── Intl.ListFormat ───────────────────────────────────────
#[test]
fn intl_listformat_conjunction() {
    assert_eq!(run_js(r#"
const lf = new Intl.ListFormat("en-US", { style: "long", type: "conjunction" });
const result = lf.format(["apples", "oranges", "bananas"]);
console.log(result.includes("and"));
"#), vec!["true"]);
}

#[test]
fn intl_listformat_disjunction() {
    assert_eq!(run_js(r#"
const lf = new Intl.ListFormat("en-US", { type: "disjunction" });
const result = lf.format(["a", "b"]);
console.log(result.includes("or"));
"#), vec!["true"]);
}

// ── Intl.RelativeTimeFormat ───────────────────────────────
#[test]
fn intl_relativetimeformat_basic() {
    assert_eq!(run_js(r#"
const rtf = new Intl.RelativeTimeFormat("en-US", { numeric: "auto" });
const result = rtf.format(-1, "day");
console.log(result);
"#), vec!["yesterday"]);
}

#[test]
fn intl_relativetimeformat_numeric() {
    assert_eq!(run_js(r#"
const rtf = new Intl.RelativeTimeFormat("en-US", { numeric: "always" });
const result = rtf.format(-3, "day");
console.log(result.includes("3"));
"#), vec!["true"]);
}

// ── Intl.Segmenter ────────────────────────────────────────
#[test]
fn intl_segmenter_words() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en-US", { granularity: "word" });
const segments = [...seg.segment("Hello world")];
const wordSegments = segments.filter(s => s.isWordLike);
console.log(wordSegments.length);
"#), vec!["2"]);
}

#[test]
fn intl_segmenter_graphemes() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en-US", { granularity: "grapheme" });
const segments = [...seg.segment("abc")];
console.log(segments.length);
"#), vec!["3"]);
}

// ── Intl.getCanonicalLocales ──────────────────────────────
#[test]
fn intl_getcanonicallocales_basic() {
    assert_eq!(run_js(r#"
const locales = Intl.getCanonicalLocales(["en-US", "fr-FR"]);
console.log(locales.length);
"#), vec!["2"]);
}

#[test]
fn intl_getcanonicallocales_single() {
    assert_eq!(run_js(r#"
const locales = Intl.getCanonicalLocales("en-US");
console.log(locales[0]);
"#), vec!["en-US"]);
}

// ── Intl.supportedValuesOf ────────────────────────────────
#[test]
fn intl_supportedvaluesof_calendar() {
    assert_eq!(run_js(r#"
const calendars = Intl.supportedValuesOf("calendar");
console.log(Array.isArray(calendars));
console.log(calendars.length > 0);
"#), vec!["true", "true"]);
}

#[test]
fn intl_supportedvaluesof_currency() {
    assert_eq!(run_js(r#"
const currencies = Intl.supportedValuesOf("currency");
console.log(currencies.includes("USD"));
"#), vec!["true"]);
}

// ── Intl locale negotiation ───────────────────────────────
#[test]
fn intl_resolvedoptions_locale() {
    assert_eq!(run_js(r#"
const fmt = new Intl.NumberFormat("en-US");
const opts = fmt.resolvedOptions();
console.log(opts.locale === "en-US" || opts.locale.startsWith("en"));
"#), vec!["true"]);
}
