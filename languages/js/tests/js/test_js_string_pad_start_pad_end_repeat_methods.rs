use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: String Padding (`padStart`, `padEnd`) & `repeat` Utility Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_string_pad_start_basic_padding() {
    let src = r#"
const str = "5";
console.log(str.padStart(3, "0"));
"#;
    assert_eq!(run_js(src), vec!["005"]);
}

#[test]
fn test_js_string_pad_end_basic_padding() {
    let src = r#"
const str = "5";
console.log(str.padEnd(3, "0"));
"#;
    assert_eq!(run_js(src), vec!["500"]);
}

#[test]
fn test_js_string_pad_start_default_pad_string_is_space() {
    let src = r#"
const str = "abc";
console.log("'" + str.padStart(5) + "'");
"#;
    assert_eq!(run_js(src), vec!["'  abc'"]);
}

#[test]
fn test_js_string_pad_end_default_pad_string_is_space() {
    let src = r#"
const str = "abc";
console.log("'" + str.padEnd(5) + "'");
"#;
    assert_eq!(run_js(src), vec!["'abc  '"]);
}

#[test]
fn test_js_string_pad_start_truncates_long_pad_string() {
    let src = r#"
const str = "abc";
console.log(str.padStart(6, "12345"));
"#;
    assert_eq!(run_js(src), vec!["123abc"]);
}

#[test]
fn test_js_string_pad_end_truncates_long_pad_string() {
    let src = r#"
const str = "abc";
console.log(str.padEnd(6, "12345"));
"#;
    assert_eq!(run_js(src), vec!["abc123"]);
}

#[test]
fn test_js_string_pad_start_target_length_less_than_string_length() {
    let src = r#"
const str = "hello";
console.log(str.padStart(3, "0"));
"#;
    assert_eq!(run_js(src), vec!["hello"]);
}

#[test]
fn test_js_string_pad_end_target_length_less_than_string_length() {
    let src = r#"
const str = "hello";
console.log(str.padEnd(3, "0"));
"#;
    assert_eq!(run_js(src), vec!["hello"]);
}

#[test]
fn test_js_string_pad_start_empty_pad_string_returns_original() {
    let src = r#"
const str = "abc";
console.log(str.padStart(10, ""));
"#;
    assert_eq!(run_js(src), vec!["abc"]);
}

#[test]
fn test_js_string_repeat_basic_multiplication() {
    let src = r#"
const str = "abc";
console.log(str.repeat(3));
"#;
    assert_eq!(run_js(src), vec!["abcabcabc"]);
}

#[test]
fn test_js_string_repeat_zero_returns_empty_string() {
    let src = r#"
const str = "abc";
console.log("'" + str.repeat(0) + "'");
"#;
    assert_eq!(run_js(src), vec!["''"]);
}

#[test]
fn test_js_string_repeat_negative_count_throws_rangeerror() {
    let src = r#"
try {
    "abc".repeat(-1);
} catch (e) {
    console.log("repeat Negative RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["repeat Negative RangeError"]);
}

#[test]
fn test_js_string_repeat_infinity_count_throws_rangeerror() {
    let src = r#"
try {
    "abc".repeat(Infinity);
} catch (e) {
    console.log("repeat Infinity RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["repeat Infinity RangeError"]);
}

#[test]
fn test_js_string_repeat_floats_truncated_to_integer() {
    let src = r#"
const str = "a";
console.log(str.repeat(3.9));
"#;
    assert_eq!(run_js(src), vec!["aaa"]);
}

#[test]
fn test_js_string_pad_start_coerces_target_length_to_integer() {
    let src = r#"
const str = "1";
console.log(str.padStart("4.8", "0"));
"#;
    assert_eq!(run_js(src), vec!["0001"]);
}

#[test]
fn test_js_string_pad_end_coerces_pad_string_to_string() {
    let src = r#"
const str = "val:";
console.log(str.padEnd(8, 123));
"#;
    assert_eq!(run_js(src), vec!["val:1231"]);
}

#[test]
fn test_js_string_pad_start_surrogate_pairs_handling() {
    let src = r#"
const emoji = "😀";
console.log(emoji.padStart(4, "🚀"));
"#;
    assert_eq!(run_js(src), vec!["🚀😀"]); // 😀 is 2 code units, target length 4 adds 2 code units ("🚀")!
}

#[test]
fn test_js_string_pad_start_null_or_undefined_this_throws_typeerror() {
    let src = r#"
try {
    String.prototype.padStart.call(null, 5, "0");
} catch (e) {
    console.log("padStart Null This TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["padStart Null This TypeError"]);
}

#[test]
fn test_js_string_repeat_null_or_undefined_this_throws_typeerror() {
    let src = r#"
try {
    String.prototype.repeat.call(undefined, 3);
} catch (e) {
    console.log("repeat Undefined This TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["repeat Undefined This TypeError"]);
}

#[test]
fn test_js_string_pad_start_property_descriptors() {
    let src = r#"
const dStart = Object.getOwnPropertyDescriptor(String.prototype, "padStart");
const dRepeat = Object.getOwnPropertyDescriptor(String.prototype, "repeat");
console.log(`${dStart.writable}:${dStart.configurable}:${dRepeat.writable}:${dRepeat.configurable}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true:true"]);
}

#[test]
fn test_js_string_repeat_nan_count_returns_empty_string() {
    let src = r#"
console.log("abc".repeat(NaN) === "");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

