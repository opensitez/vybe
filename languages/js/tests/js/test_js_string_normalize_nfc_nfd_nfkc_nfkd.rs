use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Unicode Normalization (`String.prototype.normalize()`, NFC, NFD, NFKC, NFKD)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_string_normalize_default_is_nfc() {
    let src = r#"
const str = "\u0041\u030A"; // A + combining ring above
console.log(str.normalize() === str.normalize("NFC"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_normalize_nfc_canonical_composition() {
    let src = r#"
const str = "\u0041\u030A"; // Decomposed Å
const normalized = str.normalize("NFC");
console.log(normalized + "|len=" + normalized.length + "|code=" + normalized.charCodeAt(0));
"#;
    assert_eq!(run_js(src), vec!["Å|len=1|code=197"]);
}

#[test]
fn test_js_string_normalize_nfd_canonical_decomposition() {
    let src = r#"
const str = "\u00C5"; // Composed Å
const normalized = str.normalize("NFD");
console.log(normalized.length + "|c0=" + normalized.charCodeAt(0) + "|c1=" + normalized.charCodeAt(1));
"#;
    assert_eq!(run_js(src), vec!["2|c0=65|c1=778"]);
}

#[test]
fn test_js_string_normalize_nfkc_compatibility_composition() {
    let src = r#"
const str = "\uFB01"; // 'fi' ligature
const normalized = str.normalize("NFKC");
console.log(normalized);
"#;
    assert_eq!(run_js(src), vec!["fi"]);
}

#[test]
fn test_js_string_normalize_nfkd_compatibility_decomposition() {
    let src = r#"
const str = "\u2163"; // Roman numeral IV (\u2163)
const normalized = str.normalize("NFKD");
console.log(normalized);
"#;
    assert_eq!(run_js(src), vec!["IV"]);
}

#[test]
fn test_js_string_normalize_equality_comparison() {
    let src = r#"
const s1 = "\u00C9"; // É composed
const s2 = "\u0045\u0301"; // E + combining acute
console.log((s1 === s2) + "|" + (s1.normalize() === s2.normalize()));
"#;
    assert_eq!(run_js(src), vec!["false|true"]);
}

#[test]
fn test_js_string_normalize_invalid_form_throws_rangeerror() {
    let src = r#"
try {
    "test".normalize("INVALID_FORM");
} catch (e) {
    console.log("normalize Invalid Form RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["normalize Invalid Form RangeError"]);
}

#[test]
fn test_js_string_normalize_already_normalized_returns_same_content() {
    let src = r#"
const str = "Hello World";
console.log(str.normalize("NFC") === str);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_normalize_fullwidth_characters_nfkc() {
    let src = r#"
const str = "ＡＢＣ"; // Fullwidth ABC (\uFF21\uFF22\uFF23)
console.log(str.normalize("NFKC"));
"#;
    assert_eq!(run_js(src), vec!["ABC"]);
}

#[test]
fn test_js_string_normalize_superscript_digits_nfkc() {
    let src = r#"
const str = "²³¹"; // Superscript 2, 3, 1
console.log(str.normalize("NFKC"));
"#;
    assert_eq!(run_js(src), vec!["231"]);
}

#[test]
fn test_js_string_normalize_empty_string() {
    let src = r#"
console.log("".normalize("NFC") === "");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_normalize_hangul_syllables() {
    let src = r#"
const composed = "\uAC00"; // Korean 'Ga'
const decomposed = composed.normalize("NFD");
console.log(decomposed.length + "|" + (decomposed.normalize("NFC") === composed));
"#;
    assert_eq!(run_js(src), vec!["2|true"]);
}

#[test]
fn test_js_string_normalize_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(String.prototype, "normalize");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${String.prototype.normalize.length}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:0"]);
}

#[test]
fn test_js_string_normalize_on_string_object() {
    let src = r#"
const strObj = new String("\u0041\u030A");
const normalized = strObj.normalize("NFC");
console.log(typeof normalized + "|" + normalized);
"#;
    assert_eq!(run_js(src), vec!["string|Å"]);
}

#[test]
fn test_js_string_normalize_coerces_this_to_string() {
    let src = r#"
const res = String.prototype.normalize.call(12345, "NFC");
console.log(typeof res + "|" + res);
"#;
    assert_eq!(run_js(src), vec!["string|12345"]);
}

#[test]
fn test_js_string_normalize_null_or_undefined_this_throws_typeerror() {
    let src = r#"
try {
    String.prototype.normalize.call(null);
} catch (e) {
    console.log("normalize Null This TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["normalize Null This TypeError"]);
}

#[test]
fn test_js_string_normalize_nfd_combining_character_ordering() {
    let src = r#"
const str = "e\u0301\u0327"; // e + acute + cedilla
const normalized = str.normalize("NFD");
console.log(normalized.length);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_string_normalize_fractions_nfkc() {
    let src = r#"
const str = "½"; // Vulgar fraction 1/2 (\u00BD)
console.log(str.normalize("NFKC"));
"#;
    assert_eq!(run_js(src), vec!["1⁄2"]);
}

#[test]
fn test_js_string_normalize_circled_numbers_nfkc() {
    let src = r#"
const str = "①②③"; // Circled numbers 1, 2, 3
console.log(str.normalize("NFKC"));
"#;
    assert_eq!(run_js(src), vec!["123"]);
}

#[test]
fn test_js_string_normalize_name_property() {
    let src = r#"
console.log(String.prototype.normalize.name);
"#;
    assert_eq!(run_js(src), vec!["normalize"]);
}

#[test]
fn test_js_string_normalize_form_argument_coercion() {
    let src = r#"
const formObj = { toString: () => "NFC" };
console.log("\u0041\u030A".normalize(formObj) === "Å");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

