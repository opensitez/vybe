use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Intl.Collator` String Comparison & Sensitivity Options
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_intl_collator_basic_comparison() {
    let src = r#"
const collator = new Intl.Collator("en");
console.log((collator.compare("a", "b") < 0) + "|" + (collator.compare("b", "a") > 0) + "|" + (collator.compare("a", "a") === 0));
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_intl_collator_numeric_sorting_option() {
    let src = r#"
const defaultSort = ["file1", "file10", "file2"].sort(new Intl.Collator("en").compare);
const numericSort = ["file1", "file10", "file2"].sort(new Intl.Collator("en", { numeric: true }).compare);

console.log(defaultSort.join(",") + "|" + numericSort.join(","));
"#;
    assert_eq!(run_js(src), vec!["file1,file10,file2|file1,file2,file10"]);
}

#[test]
fn test_js_intl_collator_sensitivity_base_case_accent_ignored() {
    let src = r#"
const collator = new Intl.Collator("en", { sensitivity: "base" });
console.log((collator.compare("a", "A") === 0) + "|" + (collator.compare("a", "á") === 0));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_intl_collator_sensitivity_accent_case_ignored() {
    let src = r#"
const collator = new Intl.Collator("en", { sensitivity: "accent" });
console.log((collator.compare("a", "A") === 0) + "|" + (collator.compare("a", "á") !== 0));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_intl_collator_sensitivity_variant_strict() {
    let src = r#"
const collator = new Intl.Collator("en", { sensitivity: "variant" });
console.log((collator.compare("a", "A") !== 0) + "|" + (collator.compare("a", "á") !== 0));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_intl_collator_case_first_upper() {
    let src = r#"
const upperFirst = ["a", "A", "b", "B"].sort(new Intl.Collator("en", { caseFirst: "upper" }).compare);
console.log(upperFirst.join(""));
"#;
    assert_eq!(run_js(src), vec!["AaBb"]);
}

#[test]
fn test_js_intl_collator_ignore_punctuation_option() {
    let src = r#"
const collator = new Intl.Collator("en", { ignorePunctuation: true });
console.log(collator.compare("red, envelope", "red envelope") === 0);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_collator_supported_locales_of() {
    let src = r#"
const supported = Intl.Collator.supportedLocalesOf(["en-US", "de-DE"]);
console.log(supported.includes("en-US"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_collator_resolved_options() {
    let src = r#"
const collator = new Intl.Collator("en-US", { numeric: true, sensitivity: "base" });
const opts = collator.resolvedOptions();
console.log(opts.locale + "|" + opts.numeric + "|" + opts.sensitivity);
"#;
    assert_eq!(run_js(src), vec!["en-US|true|base"]);
}

#[test]
fn test_js_intl_collator_locale_compare_string_method_integration() {
    let src = r#"
console.log("ä".localeCompare("z", "de") < 0);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_collator_array_sort_custom_comparator() {
    let src = r#"
const words = ["banana", "Apple", "cherry"];
words.sort(new Intl.Collator("en", { sensitivity: "base" }).compare);
console.log(words.join(","));
"#;
    assert_eq!(run_js(src), vec!["Apple,banana,cherry"]);
}

#[test]
fn test_js_intl_collator_invalid_sensitivity_throws_rangeerror() {
    let src = r#"
try {
    new Intl.Collator("en", { sensitivity: "invalid_sensitivity" });
} catch (e) {
    console.log("Invalid Sensitivity RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Sensitivity RangeError"]);
}

#[test]
fn test_js_intl_collator_invalid_usage_throws_rangeerror() {
    let src = r#"
try {
    new Intl.Collator("en", { usage: "invalid_usage" });
} catch (e) {
    console.log("Invalid Usage RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Usage RangeError"]);
}

#[test]
fn test_js_intl_collator_usage_search() {
    let src = r#"
const collator = new Intl.Collator("en", { usage: "search" });
console.log(collator.resolvedOptions().usage);
"#;
    assert_eq!(run_js(src), vec!["search"]);
}

#[test]
fn test_js_intl_collator_compare_identical_strings_returns_zero() {
    let src = r#"
const collator = new Intl.Collator();
console.log(collator.compare("Hello World", "Hello World"));
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_intl_collator_coercion_of_arguments_to_string() {
    let src = r#"
const collator = new Intl.Collator("en");
console.log(collator.compare(100, "100") === 0);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_collator_case_first_lower() {
    let src = r#"
const lowerFirst = ["a", "A", "b", "B"].sort(new Intl.Collator("en", { caseFirst: "lower" }).compare);
console.log(lowerFirst.join(""));
"#;
    assert_eq!(run_js(src), vec!["aAbB"]);
}

#[test]
fn test_js_intl_collator_collation_option_emoji() {
    let src = r#"
const collator = new Intl.Collator("en", { collation: "emoji" });
console.log(collator.resolvedOptions().collation);
"#;
    assert_eq!(run_js(src), vec!["emoji"]);
}

#[test]
fn test_js_intl_collator_compare_bound_method_safety() {
    let src = r#"
const compare = new Intl.Collator("en").compare;
console.log(compare("a", "b") < 0);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_intl_collator_symbol_argument_compare_throws_typeerror() {
    let src = r#"
const collator = new Intl.Collator("en");
try {
    collator.compare(Symbol("a"), "a");
} catch (e) {
    console.log("Collator Symbol Argument TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Collator Symbol Argument TypeError"]);
}
