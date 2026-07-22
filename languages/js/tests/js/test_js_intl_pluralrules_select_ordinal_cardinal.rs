use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Intl.PluralRules` Cardinal & Ordinal Categorization
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_intl_pluralrules_cardinal_english_one_other() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
console.log(`${rules.select(0)}:${rules.select(1)}:${rules.select(2)}:${rules.select(5)}`);
"#;
    assert_eq!(run_js(src), vec!["other:one:other:other"]);
}

#[test]
fn test_js_intl_pluralrules_ordinal_english_st_nd_rd_th() {
    let src = r#"
const rules = new Intl.PluralRules("en-US", { type: "ordinal" });
console.log(`${rules.select(1)}:${rules.select(2)}:${rules.select(3)}:${rules.select(4)}:${rules.select(11)}`);
"#;
    assert_eq!(run_js(src), vec!["one:two:few:other:other"]);
}

#[test]
fn test_js_intl_pluralrules_cardinal_french_zero_is_one() {
    let src = r#"
const rules = new Intl.PluralRules("fr-FR");
console.log(`${rules.select(0)}:${rules.select(1)}:${rules.select(2)}`);
"#;
    assert_eq!(run_js(src), vec!["one:one:other"]);
}

#[test]
fn test_js_intl_pluralrules_resolved_options() {
    let src = r#"
const rules = new Intl.PluralRules("en-US", { type: "ordinal" });
const opts = rules.resolvedOptions();
console.log(opts.locale + "|" + opts.type + "|" + opts.pluralCategories.join(","));
"#;
    assert_eq!(run_js(src), vec!["en-US|ordinal|one,two,few,other"]);
}

#[test]
fn test_js_intl_pluralrules_supported_locales_of() {
    let src = r#"
const supported = Intl.PluralRules.supportedLocalesOf(["en-US", "ar"]);
console.log(supported.includes("en-US"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_pluralrules_select_range_cardinal() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
console.log(rules.selectRange(0, 1) + "|" + rules.selectRange(1, 2));
"#;
    assert_eq!(run_js(src), vec!["other|other"]);
}

#[test]
fn test_js_intl_pluralrules_minimum_fraction_digits() {
    let src = r#"
const rules = new Intl.PluralRules("en-US", { minimumFractionDigits: 2 });
console.log(rules.select(1)); // 1.00 in en-US is 'other' for cardinal!
"#;
    assert_eq!(run_js(src), vec!["other"]);
}

#[test]
fn test_js_intl_pluralrules_invalid_type_option_throws_rangeerror() {
    let src = r#"
try {
    new Intl.PluralRules("en-US", { type: "invalid_type" });
} catch (e) {
    console.log("Invalid PluralRules Type RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid PluralRules Type RangeError"]);
}

#[test]
fn test_js_intl_pluralrules_select_coerces_string_number() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
console.log(rules.select("1"));
"#;
    assert_eq!(run_js(src), vec!["one"]);
}

#[test]
fn test_js_intl_pluralrules_select_negative_number() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
console.log(rules.select(-1));
"#;
    assert_eq!(run_js(src), vec!["one"]);
}

#[test]
fn test_js_intl_pluralrules_select_nan_returns_other() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
console.log(rules.select(NaN));
"#;
    assert_eq!(run_js(src), vec!["other"]);
}

#[test]
fn test_js_intl_pluralrules_select_infinity_returns_other() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
console.log(rules.select(Infinity));
"#;
    assert_eq!(run_js(src), vec!["other"]);
}

#[test]
fn test_js_intl_pluralrules_select_symbol_throws_typeerror() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
try {
    rules.select(Symbol("1"));
} catch (e) {
    console.log("PluralRules Select Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["PluralRules Select Symbol TypeError"]);
}

#[test]
fn test_js_intl_pluralrules_select_range_same_value() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
console.log(rules.selectRange(1, 1));
"#;
    assert_eq!(run_js(src), vec!["one"]);
}

#[test]
fn test_js_intl_pluralrules_cardinal_categories_property() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
console.log(rules.resolvedOptions().pluralCategories.includes("one"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_pluralrules_select_range_nan_throws_rangeerror() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
try {
    rules.selectRange(NaN, 5);
} catch (e) {
    console.log("selectRange NaN RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["selectRange NaN RangeError"]);
}

#[test]
fn test_js_intl_pluralrules_select_range_reversed_throws_rangeerror() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
try {
    rules.selectRange(5, 1);
} catch (e) {
    console.log("selectRange Reversed RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["selectRange Reversed RangeError"]);
}

#[test]
fn test_js_intl_pluralrules_locale_fallback_matching() {
    let src = r#"
const rules = new Intl.PluralRules("en-US-posix");
console.log(rules.resolvedOptions().locale.startsWith("en"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_pluralrules_maximum_fraction_digits_option() {
    let src = r#"
const rules = new Intl.PluralRules("en-US", { maximumFractionDigits: 0 });
console.log(rules.select(1.5));
"#;
    assert_eq!(run_js(src), vec!["other"]);
}

#[test]
fn test_js_intl_pluralrules_bigint_input() {
    let src = r#"
const rules = new Intl.PluralRules("en-US");
console.log(rules.select(100n));
"#;
    assert_eq!(run_js(src), vec!["other"]);
}
