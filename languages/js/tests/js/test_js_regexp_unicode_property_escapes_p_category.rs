use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: RegExp Unicode Property Escapes (`\p{...}`, `\P{...}`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_regexp_unicode_property_general_category_letter() {
    let src = r#"
const re = /\p{Letter}/gu;
console.log("a1b2C!".match(re).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b,C"]);
}

#[test]
fn test_js_regexp_unicode_property_general_category_number() {
    let src = r#"
const re = /\p{Number}/gu;
console.log("a1b2C!".match(re).join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_regexp_unicode_property_negated_P() {
    let src = r#"
const re = /\P{Number}/gu; // Everything NOT a number!
console.log("a1!".match(re).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,!"]);
}

#[test]
fn test_js_regexp_unicode_property_script_greek() {
    let src = r#"
const re = /\p{Script=Greek}/gu;
console.log("αbβ1".match(re).join(","));
"#;
    assert_eq!(run_js(src), vec!["α,β"]);
}

#[test]
fn test_js_regexp_unicode_property_script_cyrillic() {
    let src = r#"
const re = /\p{Script=Cyrillic}/gu;
console.log("Привет World".match(re).join(""));
"#;
    assert_eq!(run_js(src), vec!["Привет"]);
}

#[test]
fn test_js_regexp_unicode_property_script_hebrew() {
    let src = r#"
const re = /\p{Script=Hebrew}/gu;
console.log("שלום World".match(re).join(""));
"#;
    assert_eq!(run_js(src), vec!["שלום"]);
}

#[test]
fn test_js_regexp_unicode_property_script_arabic() {
    let src = r#"
const re = /\p{Script=Arabic}/gu;
console.log("مرحبا World".match(re).join(""));
"#;
    assert_eq!(run_js(src), vec!["مرحبا"]);
}

#[test]
fn test_js_regexp_unicode_property_binary_emoji() {
    let src = r#"
const re = /\p{Emoji}/gu;
console.log("Hi 😀 ⚽ A".match(re).join(""));
"#;
    assert_eq!(run_js(src), vec!["😀⚽"]);
}

#[test]
fn test_js_regexp_unicode_property_white_space() {
    let src = r#"
const re = /\p{White_Space}/gu;
console.log("A\tB\nC D".match(re).length);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_regexp_unicode_property_uppercase_and_lowercase() {
    let src = r#"
const upper = "aBcD".match(/\p{Uppercase}/gu).join(",");
const lower = "aBcD".match(/\p{Lowercase}/gu).join(",");
console.log(`${upper}|${lower}`);
"#;
    assert_eq!(run_js(src), vec!["B,D|a,c"]);
}

#[test]
fn test_js_regexp_unicode_property_punctuation() {
    let src = r#"
const re = /\p{Punctuation}/gu;
console.log("Hello, world!".match(re).join(","));
"#;
    assert_eq!(run_js(src), vec![",,!"]);
}

#[test]
fn test_js_regexp_unicode_property_currency_symbol() {
    let src = r#"
const re = /\p{Currency_Symbol}/gu;
console.log("$100 €20 £5 ¥1000".match(re).join(","));
"#;
    assert_eq!(run_js(src), vec!["$,€,£,¥"]);
}

#[test]
fn test_js_regexp_unicode_property_without_unicode_flag_throws_syntaxerror() {
    let src = r#"
try {
    eval("const re = /\\p{Letter}/;"); // \\p without u or v flag is a SyntaxError!
} catch (e) {
    console.log("Unicode Property Escape Without u Flag SyntaxError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Unicode Property Escape Without u Flag SyntaxError"]
    );
}

#[test]
fn test_js_regexp_unicode_property_invalid_property_name_throws_syntaxerror() {
    let src = r#"
try {
    eval("const re = /\\p{InvalidNonExistentProperty}/u;");
} catch (e) {
    console.log("Invalid Unicode Property SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Unicode Property SyntaxError"]);
}

#[test]
fn test_js_regexp_unicode_property_inside_character_class() {
    let src = r#"
const re = /[\p{Digit}\p{Letter}]/gu;
console.log("a1!".match(re).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,1"]);
}

#[test]
fn test_js_regexp_unicode_property_math_symbol() {
    let src = r#"
const re = /\p{Math}/gu;
console.log("a + b = c * d".match(re).join(","));
"#;
    assert_eq!(run_js(src), vec!["+,=,*"]);
}

#[test]
fn test_js_regexp_unicode_property_script_latin() {
    let src = r#"
const re = /\p{Script=Latin}/gu;
console.log("Hello αβ".match(re).join(""));
"#;
    assert_eq!(run_js(src), vec!["Hello"]);
}

#[test]
fn test_js_regexp_unicode_property_hex_digit() {
    let src = r#"
const re = /\p{Hex_Digit}/gu;
console.log("10AFG".match(re).join(","));
"#;
    assert_eq!(run_js(src), vec!["1,0,A,F"]);
}

#[test]
fn test_js_regexp_unicode_property_ideographic() {
    let src = r#"
const re = /\p{Ideographic}/gu;
console.log("漢字 World".match(re).join(""));
"#;
    assert_eq!(run_js(src), vec!["漢字"]);
}

#[test]
fn test_js_regexp_unicode_property_alphabetic() {
    let src = r#"
const re = /\p{Alphabetic}/gu;
console.log("a1b2c3!".match(re).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b,c"]);
}
