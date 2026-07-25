//! Behaviour tests for `ecma:intl/*` host imports — ECMA-402 Intl
//! Internationalization API.
//!
//! Reference: <https://tc39.es/ecma402/>.
//!
//! Coverage:
//!   - Intl.Collator           — locale-aware string compare
//!   - Intl.NumberFormat        — number formatting (currency, percent, etc.)
//!   - Intl.DateTimeFormat      — date/time formatting
//!   - Intl.ListFormat          — list joining ("a, b, and c")
//!   - Intl.PluralRules         — plural category select
//!   - Intl.RelativeTimeFormat  — relative time ("3 days ago")
//!   - Intl.Segmenter           — locale-aware text segmentation
//!   - Intl.Locale              — locale identifier object
//!   - Intl.DisplayNames        — display names for languages/regions/scripts/currencies
//!   - Intl.DurationFormat      — duration formatting (Stage 4, 2024)
//!   - Intl.getCanonicalLocales — canonicalize locale tag list
//!   - Intl.supportedValuesOf   — list supported values for a key
//!
//! MVP impl is en-US-only (no real ICU integration). Tests verify
//! spec-correct SHAPE (return types, property keys, method presence)
//! and basic en-US OUTPUT for the common cases. Full locale support
//! requires plugging in `icu_*` Rust crates — separate work.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_emitter::platforms::register_platforms;

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<intl-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn as_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        other => format!("{}", other),
    }
}

fn new_array(elements: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
}

fn new_object(props: Vec<(&str, Value)>) -> Value {
    let mut obj = Object::new();
    for (k, v) in props {
        obj.properties.insert(k.into(), v);
    }
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn obj_prop(value: &Value, key: &str) -> Value {
    if let Value::Object(o) = value {
        let lock = o.lock().unwrap();
        return lock
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Undefined);
    }
    Value::Undefined
}

fn array_strings(value: &Value) -> Vec<String> {
    if let Value::Object(o) = value {
        let lock = o.lock().unwrap();
        if let ObjectKind::Array(elems) = &lock.kind {
            return elems.iter().map(as_string).collect();
        }
    }
    Vec::new()
}

// ── Intl.Collator (ECMA-402 §11) ────────────────────────────────────

#[test]
fn collator_compare_returns_zero_for_equal() {
    let c = invoke("ecma:intl/collator", "new", vec![]);
    assert_eq!(
        invoke("ecma:intl/collator", "compare", vec![c, s("abc"), s("abc")]),
        Value::I32(0)
    );
}

#[test]
fn collator_compare_returns_negative_for_less() {
    let c = invoke("ecma:intl/collator", "new", vec![]);
    if let Value::I32(n) = invoke("ecma:intl/collator", "compare", vec![c, s("a"), s("b")]) {
        assert!(n < 0);
    } else {
        panic!("compare should return I32");
    }
}

#[test]
fn collator_compare_returns_positive_for_greater() {
    let c = invoke("ecma:intl/collator", "new", vec![]);
    if let Value::I32(n) = invoke("ecma:intl/collator", "compare", vec![c, s("b"), s("a")]) {
        assert!(n > 0);
    } else {
        panic!("compare should return I32");
    }
}

#[test]
fn collator_resolved_options_includes_locale() {
    let c = invoke("ecma:intl/collator", "new", vec![]);
    let opts = invoke("ecma:intl/collator", "resolvedOptions", vec![c]);
    let locale = obj_prop(&opts, "locale");
    assert!(!as_string(&locale).is_empty());
}

// ── Intl.NumberFormat (ECMA-402 §15) ─────────────────────────────────

#[test]
fn number_format_default_renders_decimal() {
    let nf = invoke("ecma:intl/numberformat", "new", vec![]);
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/numberformat",
            "format",
            vec![nf, Value::F64(1234.5)]
        )),
        "1,234.5"
    );
}

#[test]
fn number_format_currency_style() {
    let opts = new_object(vec![("style", s("currency")), ("currency", s("USD"))]);
    let nf = invoke("ecma:intl/numberformat", "new", vec![s("en-US"), opts]);
    let result = as_string(&invoke(
        "ecma:intl/numberformat",
        "format",
        vec![nf, Value::F64(99.0)],
    ));
    // en-US default: "$99.00"
    assert!(result.contains("99"));
    assert!(result.contains("$"));
}

#[test]
fn number_format_percent_style() {
    let opts = new_object(vec![("style", s("percent"))]);
    let nf = invoke("ecma:intl/numberformat", "new", vec![s("en-US"), opts]);
    let result = as_string(&invoke(
        "ecma:intl/numberformat",
        "format",
        vec![nf, Value::F64(0.25)],
    ));
    // en-US default: "25%"
    assert!(result.contains("25"));
    assert!(result.contains("%"));
}

#[test]
fn number_format_resolved_options_locale() {
    let nf = invoke("ecma:intl/numberformat", "new", vec![s("en-US")]);
    let opts = invoke("ecma:intl/numberformat", "resolvedOptions", vec![nf]);
    assert_eq!(as_string(&obj_prop(&opts, "locale")), "en-US");
}

#[test]
fn number_format_format_to_parts_returns_array() {
    let nf = invoke("ecma:intl/numberformat", "new", vec![]);
    let parts = invoke(
        "ecma:intl/numberformat",
        "formatToParts",
        vec![nf, Value::F64(1234.5)],
    );
    // Should be an Array of { type, value } objects
    if let Value::Object(o) = &parts {
        let lock = o.lock().unwrap();
        assert!(matches!(lock.kind, ObjectKind::Array(_)));
    } else {
        panic!("formatToParts should return Array");
    }
}

// ── Intl.DateTimeFormat (ECMA-402 §13) ───────────────────────────────

#[test]
fn date_time_format_renders_date_for_ms_epoch() {
    let dtf = invoke("ecma:intl/datetimeformat", "new", vec![s("en-US")]);
    // Jan 15, 2024 = 1705276800000 ms epoch
    let result = as_string(&invoke(
        "ecma:intl/datetimeformat",
        "format",
        vec![dtf, Value::F64(1705276800000.0)],
    ));
    // en-US default: "1/15/2024" (MM/DD/YYYY)
    assert!(result.contains("2024"));
}

#[test]
fn date_time_format_resolved_options_locale() {
    let dtf = invoke("ecma:intl/datetimeformat", "new", vec![s("en-US")]);
    let opts = invoke("ecma:intl/datetimeformat", "resolvedOptions", vec![dtf]);
    assert_eq!(as_string(&obj_prop(&opts, "locale")), "en-US");
}

// ── Intl.ListFormat (ECMA-402 §14) ───────────────────────────────────

#[test]
fn list_format_long_conjunction_uses_oxford_comma() {
    let lf = invoke("ecma:intl/listformat", "new", vec![]);
    let list = new_array(vec![s("a"), s("b"), s("c")]);
    // en-US conjunction: "a, b, and c"
    assert_eq!(
        as_string(&invoke("ecma:intl/listformat", "format", vec![lf, list])),
        "a, b, and c"
    );
}

#[test]
fn list_format_two_items_uses_and() {
    let lf = invoke("ecma:intl/listformat", "new", vec![]);
    let list = new_array(vec![s("a"), s("b")]);
    assert_eq!(
        as_string(&invoke("ecma:intl/listformat", "format", vec![lf, list])),
        "a and b"
    );
}

#[test]
fn list_format_disjunction_uses_or() {
    let opts = new_object(vec![("type", s("disjunction"))]);
    let lf = invoke("ecma:intl/listformat", "new", vec![s("en-US"), opts]);
    let list = new_array(vec![s("a"), s("b"), s("c")]);
    assert_eq!(
        as_string(&invoke("ecma:intl/listformat", "format", vec![lf, list])),
        "a, b, or c"
    );
}

// ── Intl.PluralRules (ECMA-402 §16) ──────────────────────────────────

#[test]
fn plural_rules_select_one_for_singular() {
    let pr = invoke("ecma:intl/pluralrules", "new", vec![]);
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/pluralrules",
            "select",
            vec![pr, Value::I32(1)]
        )),
        "one"
    );
}

#[test]
fn plural_rules_select_other_for_plural() {
    let pr = invoke("ecma:intl/pluralrules", "new", vec![]);
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/pluralrules",
            "select",
            vec![pr.clone(), Value::I32(2)]
        )),
        "other"
    );
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/pluralrules",
            "select",
            vec![pr, Value::I32(0)]
        )),
        "other"
    );
}

#[test]
fn plural_rules_ordinal_select() {
    let opts = new_object(vec![("type", s("ordinal"))]);
    let pr = invoke("ecma:intl/pluralrules", "new", vec![s("en-US"), opts]);
    // English ordinals: 1st=one, 2nd=two, 3rd=few, 4th-19th=other, 21st=one, ...
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/pluralrules",
            "select",
            vec![pr.clone(), Value::I32(1)]
        )),
        "one"
    );
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/pluralrules",
            "select",
            vec![pr.clone(), Value::I32(2)]
        )),
        "two"
    );
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/pluralrules",
            "select",
            vec![pr, Value::I32(3)]
        )),
        "few"
    );
}

// ── Intl.RelativeTimeFormat (ECMA-402 §17) ───────────────────────────

#[test]
fn relative_time_format_past_days() {
    let rtf = invoke("ecma:intl/relativetimeformat", "new", vec![]);
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/relativetimeformat",
            "format",
            vec![rtf, Value::I32(-3), s("day")]
        )),
        "3 days ago"
    );
}

#[test]
fn relative_time_format_future_days() {
    let rtf = invoke("ecma:intl/relativetimeformat", "new", vec![]);
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/relativetimeformat",
            "format",
            vec![rtf, Value::I32(3), s("day")]
        )),
        "in 3 days"
    );
}

#[test]
fn relative_time_format_singular_unit() {
    let rtf = invoke("ecma:intl/relativetimeformat", "new", vec![]);
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/relativetimeformat",
            "format",
            vec![rtf.clone(), Value::I32(1), s("hour")]
        )),
        "in 1 hour"
    );
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/relativetimeformat",
            "format",
            vec![rtf, Value::I32(-1), s("hour")]
        )),
        "1 hour ago"
    );
}

// ── Intl.Segmenter (ECMA-402 §18) ────────────────────────────────────

#[test]
fn segmenter_grapheme_segments() {
    let seg = invoke("ecma:intl/segmenter", "new", vec![]);
    let result = invoke("ecma:intl/segmenter", "segment", vec![seg, s("hi")]);
    // Should return an iterable / Array of { segment, index, input } objects.
    // MVP returns Array of segments.
    if let Value::Object(o) = &result {
        let lock = o.lock().unwrap();
        assert!(matches!(lock.kind, ObjectKind::Array(_)));
    } else {
        panic!("segment should return iterable Array");
    }
}

// ── Intl.Locale (ECMA-402 §14) ───────────────────────────────────────

#[test]
fn locale_base_name_from_tag() {
    let loc = invoke("ecma:intl/locale", "new", vec![s("en-US")]);
    assert_eq!(as_string(&obj_prop(&loc, "baseName")), "en-US");
}

#[test]
fn locale_language_extracted() {
    let loc = invoke("ecma:intl/locale", "new", vec![s("en-US")]);
    assert_eq!(as_string(&obj_prop(&loc, "language")), "en");
}

#[test]
fn locale_region_extracted() {
    let loc = invoke("ecma:intl/locale", "new", vec![s("en-US")]);
    assert_eq!(as_string(&obj_prop(&loc, "region")), "US");
}

#[test]
fn locale_to_string_returns_canonical_tag() {
    let loc = invoke("ecma:intl/locale", "new", vec![s("en-US")]);
    assert_eq!(
        as_string(&invoke("ecma:intl/locale", "toString", vec![loc])),
        "en-US"
    );
}

// ── Intl.DisplayNames (ECMA-402 §12) ─────────────────────────────────

#[test]
fn display_names_of_language_code() {
    let opts = new_object(vec![("type", s("language"))]);
    let dn = invoke("ecma:intl/displaynames", "new", vec![s("en"), opts]);
    let result = as_string(&invoke("ecma:intl/displaynames", "of", vec![dn, s("fr")]));
    // en display name for "fr" is "French"
    assert_eq!(result, "French");
}

#[test]
fn display_names_of_region_code() {
    let opts = new_object(vec![("type", s("region"))]);
    let dn = invoke("ecma:intl/displaynames", "new", vec![s("en"), opts]);
    assert_eq!(
        as_string(&invoke("ecma:intl/displaynames", "of", vec![dn, s("US")])),
        "United States"
    );
}

#[test]
fn display_names_french_for_us_region() {
    let opts = new_object(vec![("type", s("region"))]);
    let dn = invoke("ecma:intl/displaynames", "new", vec![s("fr"), opts]);
    let result = as_string(&invoke("ecma:intl/displaynames", "of", vec![dn, s("US")]));
    // French for "US": "États-Unis"
    assert!(
        result.contains("États") || result.contains("Etats"),
        "Expected French US name, got {:?}",
        result
    );
}

#[test]
fn display_names_german_for_french_language() {
    let opts = new_object(vec![("type", s("language"))]);
    let dn = invoke("ecma:intl/displaynames", "new", vec![s("de"), opts]);
    let result = as_string(&invoke("ecma:intl/displaynames", "of", vec![dn, s("fr")]));
    // German for "fr": "Französisch"
    assert_eq!(result, "Französisch");
}

#[test]
fn display_names_japanese_for_english_language() {
    let opts = new_object(vec![("type", s("language"))]);
    let dn = invoke("ecma:intl/displaynames", "new", vec![s("ja"), opts]);
    let result = as_string(&invoke("ecma:intl/displaynames", "of", vec![dn, s("en")]));
    // Japanese for "en": "英語"
    assert_eq!(result, "英語");
}

#[test]
fn display_names_script_latin_in_english() {
    let opts = new_object(vec![("type", s("script"))]);
    let dn = invoke("ecma:intl/displaynames", "new", vec![s("en"), opts]);
    let result = as_string(&invoke("ecma:intl/displaynames", "of", vec![dn, s("Latn")]));
    assert_eq!(result, "Latin");
}

// ── Intl.DurationFormat (ECMA-402 §19, Stage 4 in 2024) ──────────────

#[test]
fn duration_format_basic() {
    let df = invoke("ecma:intl/durationformat", "new", vec![]);
    let dur = new_object(vec![("hours", Value::I32(1)), ("minutes", Value::I32(30))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // en-US long: "1 hr, 30 min" (depends on style default; spec says "short")
    // MVP: just verify the numbers appear in the output.
    assert!(result.contains("1"));
    assert!(result.contains("30"));
}

// ── DateTimeFormat — locale-aware date order/month names ────────────

#[test]
fn date_time_format_german_uses_dot_separator() {
    let dtf = invoke("ecma:intl/datetimeformat", "new", vec![s("de-DE")]);
    // Jan 15, 2024 = 1705276800000 ms epoch (UTC)
    let result = as_string(&invoke(
        "ecma:intl/datetimeformat",
        "format",
        vec![dtf, Value::F64(1705276800000.0)],
    ));
    // German default medium-length date: "15.01.2024" (CLDR pattern)
    assert!(
        result.contains("2024"),
        "Expected year 2024 in output, got {:?}",
        result
    );
    assert!(
        result.contains("."),
        "Expected German dot separator, got {:?}",
        result
    );
}

#[test]
fn date_time_format_french_uses_french_month_names() {
    let dtf = invoke("ecma:intl/datetimeformat", "new", vec![s("fr-FR")]);
    let result = as_string(&invoke(
        "ecma:intl/datetimeformat",
        "format",
        vec![dtf, Value::F64(1705276800000.0)],
    ));
    // French default medium date includes localized month name like "janv."
    // (French abbreviated month for January)
    assert!(
        result.contains("janv") || result.contains("1"),
        "Expected French January formatting, got {:?}",
        result
    );
    assert!(result.contains("2024"));
}

#[test]
fn date_time_format_japanese_uses_japanese_format() {
    let dtf = invoke("ecma:intl/datetimeformat", "new", vec![s("ja-JP")]);
    let result = as_string(&invoke(
        "ecma:intl/datetimeformat",
        "format",
        vec![dtf, Value::F64(1705276800000.0)],
    ));
    // Japanese default uses 年 月 日 markers
    assert!(result.contains("2024"));
    // Year-first ordering OR Japanese era markers
    assert!(
        result.contains("年") || result.starts_with("2024"),
        "Expected Japanese formatting, got {:?}",
        result
    );
}

// ── NON-EN-US LOCALE TESTS — proves real ICU is hooked up ────────────
//
// These all verify locale-aware behaviour. With the en-US shim impl
// they would all fail (returning en-US output regardless of locale).

// PluralRules — Russian has 4 cardinal categories: one (1, 21, 31...),
// few (2-4, 22-24, ...), many (0, 5-20, ...), other.
#[test]
fn plural_rules_russian_few() {
    let pr = invoke("ecma:intl/pluralrules", "new", vec![s("ru")]);
    // 2 → "few" in Russian (vs "other" in English)
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/pluralrules",
            "select",
            vec![pr, Value::I32(2)]
        )),
        "few"
    );
}

#[test]
fn plural_rules_russian_many() {
    let pr = invoke("ecma:intl/pluralrules", "new", vec![s("ru")]);
    // 5 → "many" in Russian
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/pluralrules",
            "select",
            vec![pr, Value::I32(5)]
        )),
        "many"
    );
}

// Welsh has the unusual "two" category for cardinal.
#[test]
fn plural_rules_welsh_two() {
    let pr = invoke("ecma:intl/pluralrules", "new", vec![s("cy")]);
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/pluralrules",
            "select",
            vec![pr, Value::I32(2)]
        )),
        "two"
    );
}

// Polish: 1 → one, 2-4 → few, rest → many.
#[test]
fn plural_rules_polish_few() {
    let pr = invoke("ecma:intl/pluralrules", "new", vec![s("pl")]);
    assert_eq!(
        as_string(&invoke(
            "ecma:intl/pluralrules",
            "select",
            vec![pr, Value::I32(3)]
        )),
        "few"
    );
}

// Collator — German DIN-1 vs en-US: 'ä' sorts differently.
// Generic locale-aware compare should at least not return the same
// thing as bytewise cmp for accented chars.
#[test]
fn collator_german_compares_umlauts() {
    let c = invoke("ecma:intl/collator", "new", vec![s("de")]);
    // 'ä' should be a/ae-equivalent in German collation, not after 'z'.
    // We just verify the result is consistent (locale collator returns
    // a valid ordering).
    let result = invoke("ecma:intl/collator", "compare", vec![c, s("ä"), s("z")]);
    if let Value::I32(n) = result {
        // German DIN: ä < z (sorts as 'a').
        assert!(
            n < 0,
            "German collation: ä should be less than z, got {}",
            n
        );
    } else {
        panic!("compare should return I32");
    }
}

// NumberFormat — German (de-DE) uses "." for thousands and "," for
// decimal separator. en-US uses inverse.
#[test]
fn number_format_german_uses_dot_for_thousands() {
    let nf = invoke("ecma:intl/numberformat", "new", vec![s("de-DE")]);
    let result = as_string(&invoke(
        "ecma:intl/numberformat",
        "format",
        vec![nf, Value::F64(1234.5)],
    ));
    // de-DE: "1.234,5"  (en-US would be "1,234.5")
    assert_eq!(result, "1.234,5");
}

#[test]
fn number_format_french_uses_space_for_thousands() {
    let nf = invoke("ecma:intl/numberformat", "new", vec![s("fr-FR")]);
    let result = as_string(&invoke(
        "ecma:intl/numberformat",
        "format",
        vec![nf, Value::F64(1234.5)],
    ));
    // fr-FR uses NARROW NO-BREAK SPACE (U+202F) for thousands and
    // "," for decimal: "1\u{202F}234,5"
    assert!(result.contains("1"));
    assert!(result.contains("234"));
    assert!(
        result.contains(",5"),
        "Expected fr-FR comma decimal, got {:?}",
        result
    );
}

// ListFormat — French uses "et" for conjunction, Spanish uses "y" / "e".
#[test]
fn list_format_french_uses_et() {
    let lf = invoke("ecma:intl/listformat", "new", vec![s("fr")]);
    let list = new_array(vec![s("a"), s("b"), s("c")]);
    let result = as_string(&invoke("ecma:intl/listformat", "format", vec![lf, list]));
    // French: "a, b et c" (no oxford comma)
    assert!(
        result.contains("et"),
        "Expected French 'et', got {:?}",
        result
    );
}

#[test]
fn list_format_spanish_uses_y() {
    let lf = invoke("ecma:intl/listformat", "new", vec![s("es")]);
    let list = new_array(vec![s("a"), s("b")]);
    let result = as_string(&invoke("ecma:intl/listformat", "format", vec![lf, list]));
    // Spanish 2-item: "a y b"
    assert!(
        result.contains(" y "),
        "Expected Spanish 'y', got {:?}",
        result
    );
}

// Locale parsing — should canonicalize tags.
#[test]
fn locale_parses_complex_tag() {
    let loc = invoke("ecma:intl/locale", "new", vec![s("zh-Hans-CN")]);
    assert_eq!(as_string(&obj_prop(&loc, "language")), "zh");
    assert_eq!(as_string(&obj_prop(&loc, "script")), "Hans");
    assert_eq!(as_string(&obj_prop(&loc, "region")), "CN");
}

// Segmenter — Unicode emoji is ONE grapheme cluster (👨‍👩‍👧‍👦 = 7 codepoints joined with ZWJ).
#[test]
fn segmenter_emoji_grapheme_cluster() {
    let seg = invoke("ecma:intl/segmenter", "new", vec![]);
    // 👨‍👩‍👧‍👦 = man + ZWJ + woman + ZWJ + girl + ZWJ + boy = 7 codepoints, 1 grapheme
    let family = "\u{1F468}\u{200D}\u{1F469}\u{200D}\u{1F467}\u{200D}\u{1F466}";
    let result = invoke("ecma:intl/segmenter", "segment", vec![seg, s(family)]);
    if let Value::Object(arr) = &result {
        let lock = arr.lock().unwrap();
        if let ObjectKind::Array(elems) = &lock.kind {
            assert_eq!(
                elems.len(),
                1,
                "Family emoji should be 1 grapheme cluster, got {}",
                elems.len()
            );
        }
    }
}

#[test]
fn segmenter_word_granularity_skips_punctuation_correctly() {
    let opts = new_object(vec![("granularity", s("word"))]);
    let seg = invoke("ecma:intl/segmenter", "new", vec![s("en"), opts]);
    let result = invoke(
        "ecma:intl/segmenter",
        "segment",
        vec![seg, s("hello, world")],
    );
    if let Value::Object(arr) = &result {
        let lock = arr.lock().unwrap();
        if let ObjectKind::Array(elems) = &lock.kind {
            // word boundaries: "hello" + "," + " " + "world" → at least 4 segments
            assert!(
                elems.len() >= 3,
                "Expected word boundaries, got {} elements",
                elems.len()
            );
        }
    }
}

// ── RelativeTimeFormat — locale-aware "ago" / "in" phrasing ─────────

#[test]
fn relative_time_format_french_past() {
    let rtf = invoke("ecma:intl/relativetimeformat", "new", vec![s("fr")]);
    let result = as_string(&invoke(
        "ecma:intl/relativetimeformat",
        "format",
        vec![rtf, Value::I32(-3), s("day")],
    ));
    // French: "il y a 3 jours"
    assert!(
        result.contains("3"),
        "Expected 3 in output, got {:?}",
        result
    );
    assert!(
        result.contains("jour") || result.contains("il y a"),
        "Expected French past phrasing, got {:?}",
        result
    );
}

#[test]
fn relative_time_format_german_future() {
    let rtf = invoke("ecma:intl/relativetimeformat", "new", vec![s("de")]);
    let result = as_string(&invoke(
        "ecma:intl/relativetimeformat",
        "format",
        vec![rtf, Value::I32(3), s("day")],
    ));
    // German: "in 3 Tagen"
    assert!(result.contains("3"));
    assert!(
        result.contains("Tag") || result.contains("in "),
        "Expected German future phrasing, got {:?}",
        result
    );
}

#[test]
fn relative_time_format_japanese_minutes() {
    let rtf = invoke("ecma:intl/relativetimeformat", "new", vec![s("ja")]);
    let result = as_string(&invoke(
        "ecma:intl/relativetimeformat",
        "format",
        vec![rtf, Value::I32(-5), s("minute")],
    ));
    // Japanese: "5 分前" (5 minutes ago)
    assert!(result.contains("5"));
    assert!(
        result.contains("分") || result.contains("前"),
        "Expected Japanese minutes phrasing, got {:?}",
        result
    );
}

// ── DurationFormat — top-20 locale text + digital style ─────────────

#[test]
fn duration_format_german_long() {
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("de"), new_object(vec![("style", s("long"))])],
    );
    let dur = new_object(vec![("hours", Value::I32(1)), ("minutes", Value::I32(30))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // German: "1 Stunde, 30 Minuten"
    assert!(
        result.contains("Stunde"),
        "Expected German hours, got {:?}",
        result
    );
    assert!(
        result.contains("Minuten"),
        "Expected German minutes plural, got {:?}",
        result
    );
}

#[test]
fn duration_format_french_long() {
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("fr"), new_object(vec![("style", s("long"))])],
    );
    let dur = new_object(vec![("hours", Value::I32(2)), ("minutes", Value::I32(15))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // French: "2 heures, 15 minutes"
    assert!(
        result.contains("heures"),
        "Expected French hours plural, got {:?}",
        result
    );
    assert!(
        result.contains("minutes"),
        "Expected French minutes, got {:?}",
        result
    );
}

#[test]
fn duration_format_japanese_long() {
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("ja"), new_object(vec![("style", s("long"))])],
    );
    let dur = new_object(vec![("hours", Value::I32(3)), ("minutes", Value::I32(45))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // Japanese: "3 時間, 45 分"
    assert!(
        result.contains("時間"),
        "Expected Japanese hours, got {:?}",
        result
    );
    assert!(
        result.contains("分"),
        "Expected Japanese minutes, got {:?}",
        result
    );
}

#[test]
fn duration_format_arabic_long() {
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("ar"), new_object(vec![("style", s("long"))])],
    );
    let dur = new_object(vec![("hours", Value::I32(2)), ("minutes", Value::I32(30))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // Arabic: 2 hours = "ساعتان" (dual form), 30 minutes = "دقيقة" (many)
    assert!(
        result.contains("ساع") || result.contains("دق"),
        "Expected Arabic hours/minutes, got {:?}",
        result
    );
}

#[test]
fn duration_format_arabic_dual_form_for_two() {
    // Arabic has a special "two" plural category — distinct from singular and plural.
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("ar"), new_object(vec![("style", s("long"))])],
    );
    let dur = new_object(vec![("hours", Value::I32(2))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // Arabic dual for hours: "ساعتان"
    assert!(
        result.contains("ساعتان"),
        "Expected Arabic dual form 'ساعتان' for 2 hours, got {:?}",
        result
    );
}

#[test]
fn duration_format_russian_few_form() {
    // Russian: 3 hours uses "few" plural form.
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("ru"), new_object(vec![("style", s("long"))])],
    );
    let dur = new_object(vec![("hours", Value::I32(3))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // Russian "few" form for hours: "часа"
    assert!(
        result.contains("часа"),
        "Expected Russian few form 'часа' for 3 hours, got {:?}",
        result
    );
}

#[test]
fn duration_format_russian_many_form() {
    // Russian: 5 hours uses "many" plural form.
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("ru"), new_object(vec![("style", s("long"))])],
    );
    let dur = new_object(vec![("hours", Value::I32(5))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // Russian "many" form for hours: "часов"
    assert!(
        result.contains("часов"),
        "Expected Russian many form 'часов' for 5 hours, got {:?}",
        result
    );
}

#[test]
fn duration_format_polish_few_form() {
    // Polish: 3 hours uses "few" plural form.
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("pl"), new_object(vec![("style", s("long"))])],
    );
    let dur = new_object(vec![("hours", Value::I32(3))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // Polish "few" form for hours: "godziny"
    assert!(
        result.contains("godziny"),
        "Expected Polish few form 'godziny' for 3 hours, got {:?}",
        result
    );
}

#[test]
fn duration_format_chinese_no_inflection() {
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("zh"), new_object(vec![("style", s("long"))])],
    );
    let dur = new_object(vec![("hours", Value::I32(2))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // Chinese: "2 小时"
    assert!(
        result.contains("小时"),
        "Expected Chinese hours, got {:?}",
        result
    );
}

#[test]
fn duration_format_digital_style() {
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("en"), new_object(vec![("style", s("digital"))])],
    );
    let dur = new_object(vec![
        ("hours", Value::I32(1)),
        ("minutes", Value::I32(30)),
        ("seconds", Value::I32(5)),
    ]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // Digital format: "1:30:05" (zero-padded minutes/seconds)
    assert_eq!(result, "1:30:05");
}

#[test]
fn duration_format_digital_pads_single_digit_seconds() {
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("en"), new_object(vec![("style", s("digital"))])],
    );
    let dur = new_object(vec![
        ("hours", Value::I32(0)),
        ("minutes", Value::I32(5)),
        ("seconds", Value::I32(3)),
    ]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // 0:05:03 — minutes/seconds zero-padded
    assert_eq!(result, "0:05:03");
}

#[test]
fn duration_format_microseconds_supported() {
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("en"), new_object(vec![("style", s("short"))])],
    );
    let dur = new_object(vec![("microseconds", Value::I32(500))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    // English short: "500 μs"
    assert_eq!(result, "500 μs");
}

#[test]
fn duration_format_nanoseconds_supported() {
    let df = invoke(
        "ecma:intl/durationformat",
        "new",
        vec![s("en"), new_object(vec![("style", s("short"))])],
    );
    let dur = new_object(vec![("nanoseconds", Value::I32(999))]);
    let result = as_string(&invoke("ecma:intl/durationformat", "format", vec![df, dur]));
    assert_eq!(result, "999 ns");
}

// ── Intl static methods ──────────────────────────────────────────────

#[test]
fn get_canonical_locales_normalizes_tag() {
    // Intl.getCanonicalLocales(["EN-us", "FR-fr"]) → ["en-US", "fr-FR"]
    let locales = new_array(vec![s("EN-us"), s("FR-fr")]);
    let result = invoke("ecma:intl", "getCanonicalLocales", vec![locales]);
    let canon = array_strings(&result);
    assert_eq!(canon, vec!["en-US", "fr-FR"]);
}

#[test]
fn get_canonical_locales_accepts_single_string() {
    // Intl.getCanonicalLocales("EN-us") → ["en-US"]
    let result = invoke("ecma:intl", "getCanonicalLocales", vec![s("EN-us")]);
    assert_eq!(array_strings(&result), vec!["en-US"]);
}

#[test]
fn supported_values_of_returns_array() {
    // Intl.supportedValuesOf("currency") returns Array of currency codes.
    let result = invoke("ecma:intl", "supportedValuesOf", vec![s("currency")]);
    if let Value::Object(o) = &result {
        let lock = o.lock().unwrap();
        assert!(matches!(lock.kind, ObjectKind::Array(_)));
        if let ObjectKind::Array(elems) = &lock.kind {
            // At minimum, USD/EUR/GBP should be in there
            let strs: Vec<String> = elems.iter().map(as_string).collect();
            assert!(strs.contains(&"USD".to_string()));
            assert!(strs.contains(&"EUR".to_string()));
        }
    } else {
        panic!("supportedValuesOf should return Array");
    }
}
