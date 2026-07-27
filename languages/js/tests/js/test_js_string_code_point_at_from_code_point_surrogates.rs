use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Unicode Code Points (`codePointAt`, `String.fromCodePoint`) & UTF-16 Surrogates
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_string_code_point_at_bmp_characters() {
    let src = r#"
const str = "ABC";
console.log(`${str.codePointAt(0)}:${str.codePointAt(1)}:${str.codePointAt(2)}`);
"#;
    assert_eq!(run_js(src), vec!["65:66:67"]);
}

#[test]
fn test_js_string_code_point_at_astral_surrogate_pairs() {
    let src = r#"
const emoji = "😀"; // U+1F600 (represented as high/low surrogate pair)
console.log(`${emoji.length}:${emoji.codePointAt(0)}:${emoji.charCodeAt(0)}`);
"#;
    assert_eq!(run_js(src), vec!["2:128512:55357"]);
}

#[test]
fn test_js_string_code_point_at_low_surrogate_position() {
    let src = r#"
const emoji = "😀";
console.log(emoji.codePointAt(1)); // At trailing surrogate position, codePointAt returns trail surrogate char code!
"#;
    assert_eq!(run_js(src), vec!["56832"]);
}

#[test]
fn test_js_string_from_code_point_multiple_args() {
    let src = r#"
console.log(String.fromCodePoint(65, 66, 128512));
"#;
    assert_eq!(run_js(src), vec!["AB😀"]);
}

#[test]
fn test_js_string_from_code_point_invalid_code_point_throws_rangeerror() {
    let src = r#"
try {
    String.fromCodePoint(0x110000); // Beyond U+10FFFF max code point!
} catch (e) {
    console.log("fromCodePoint RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["fromCodePoint RangeError"]);
}

#[test]
fn test_js_string_from_code_point_negative_code_point_throws_rangeerror() {
    let src = r#"
try {
    String.fromCodePoint(-1);
} catch (e) {
    console.log("fromCodePoint Negative RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["fromCodePoint Negative RangeError"]);
}

#[test]
fn test_js_string_from_code_point_minus_zero_is_zero() {
    let src = r#"
const c = String.fromCodePoint(-0);
console.log(`${c.length}:${c.charCodeAt(0)}`);
"#;
    assert_eq!(run_js(src), vec!["1:0"]);
}

#[test]
fn test_js_string_code_point_at_index_out_of_bounds() {
    let src = r#"
const str = "A";
console.log(str.codePointAt(5) === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_code_point_at_negative_index_returns_undefined() {
    let src = r#"
const str = "A";
console.log(str.codePointAt(-1) === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_for_of_loop_iterates_by_code_points() {
    let src = r#"
const str = "A😀B";
const codePoints = [];
for (const char of str) {
    codePoints.push(char.codePointAt(0));
}
console.log(codePoints.join(","));
"#;
    assert_eq!(run_js(src), vec!["65,128512,66"]); // for...of iterates over full code points, not UTF-16 code units!
}

#[test]
fn test_js_string_spread_operator_splits_by_code_points() {
    let src = r#"
const str = "A😀B";
const chars = [...str];
console.log(chars.length + "|" + chars[1]);
"#;
    assert_eq!(run_js(src), vec!["3|😀"]);
}

#[test]
fn test_js_string_char_code_at_vs_code_point_at() {
    let src = r#"
const str = "🎉"; // U+1F389 (127881)
console.log(`${str.charCodeAt(0)} vs ${str.codePointAt(0)}`);
"#;
    assert_eq!(run_js(src), vec!["55357 vs 127881"]);
}

#[test]
fn test_js_string_from_char_code_vs_from_code_point() {
    let src = r#"
console.log(String.fromCodePoint(127881) === String.fromCharCode(55357, 56199));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_code_point_at_coercion_of_index() {
    let src = r#"
const str = "ABC";
console.log(str.codePointAt("1.9")); // Coerces index "1.9" to 1
"#;
    assert_eq!(run_js(src), vec!["66"]);
}

#[test]
fn test_js_string_from_code_point_coercion_of_args() {
    let src = r#"
console.log(String.fromCodePoint("65", true ? 66 : 0));
"#;
    assert_eq!(run_js(src), vec!["AB"]);
}

#[test]
fn test_js_string_from_code_point_nan_throws_rangeerror() {
    let src = r#"
try {
    String.fromCodePoint(NaN);
} catch (e) {
    console.log("fromCodePoint NaN RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["fromCodePoint NaN RangeError"]);
}

#[test]
fn test_js_string_code_point_at_unpaired_surrogate() {
    let src = r#"
const loneHigh = "\uD83D"; // Unpaired lead surrogate
console.log(loneHigh.codePointAt(0));
"#;
    assert_eq!(run_js(src), vec!["55357"]);
}

#[test]
fn test_js_string_code_point_at_hex_formatting() {
    let src = r#"
const emoji = "🚀";
console.log("U+" + emoji.codePointAt(0).toString(16).toUpperCase());
"#;
    assert_eq!(run_js(src), vec!["U+1F680"]);
}

#[test]
fn test_js_string_from_code_point_zero() {
    let src = r#"
const nullChar = String.fromCodePoint(0);
console.log(nullChar.charCodeAt(0) + "|" + nullChar.length);
"#;
    assert_eq!(run_js(src), vec!["0|1"]);
}

#[test]
fn test_js_string_from_code_point_max_valid_code_point() {
    let src = r#"
const maxChar = String.fromCodePoint(0x10FFFF);
console.log(maxChar.length);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_string_array_from_code_point_mapping() {
    let src = r#"
const codes = [72, 69, 76, 76, 79];
const str = String.fromCodePoint(...codes);
console.log(str);
"#;
    assert_eq!(run_js(src), vec!["HELLO"]);
}
