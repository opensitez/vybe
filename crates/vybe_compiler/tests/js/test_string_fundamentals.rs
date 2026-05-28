/// String interpolation, template expressions, multiline, escape sequences,
/// raw strings, unicode escapes, tagged template interactions.

use super::helpers::run_js;

// ── escape sequences in strings ───────────────────────────────────────────────

#[test]
fn escape_sequence_newline_tab() {
    assert_eq!(run_js(r#"
const s = "line1\nline2\ttab";
const parts = s.split("\n");
console.log(parts[0]);
console.log(parts[1].startsWith("line2\t"));
"#), vec!["line1", "true"]);
}

#[test]
fn escape_sequence_unicode_code_point() {
    assert_eq!(run_js(r#"
console.log("\u0041");   // A
console.log("\u{1F600}"); // 😀
console.log("\u{0041}"); // A via curly syntax
"#), vec!["A", "😀", "A"]);
}

#[test]
fn escape_sequence_hex() {
    assert_eq!(run_js(r#"
console.log("\x41"); // A
console.log("\x61"); // a
"#), vec!["A", "a"]);
}

#[test]
fn null_character_escape() {
    assert_eq!(run_js(r#"
const s = "a\0b";
console.log(s.length);
console.log(s.charCodeAt(1));
"#), vec!["3", "0"]);
}

// ── string concatenation ──────────────────────────────────────────────────────

#[test]
fn string_concat_method() {
    assert_eq!(run_js(r#"
const s = "Hello".concat(", ", "World", "!");
console.log(s);
"#), vec!["Hello, World!"]);
}

#[test]
fn string_plus_coerces_non_strings() {
    assert_eq!(run_js(r#"
console.log("value: " + 42);
console.log("flag: " + true);
console.log("nothing: " + null);
"#), vec!["value: 42", "flag: true", "nothing: null"]);
}

// ── string iteration ──────────────────────────────────────────────────────────

#[test]
fn string_for_of_yields_characters() {
    assert_eq!(run_js(r#"
const chars = [];
for (const c of "abc") chars.push(c);
console.log(chars.join("-"));
"#), vec!["a-b-c"]);
}

#[test]
fn spread_string_into_array() {
    assert_eq!(run_js(r#"
const arr = [..."hello"];
console.log(arr.join(","));
"#), vec!["h,e,l,l,o"]);
}

// ── string searching ──────────────────────────────────────────────────────────

#[test]
fn indexof_and_lastindexof() {
    assert_eq!(run_js(r#"
const s = "abcabc";
console.log(s.indexOf("b"));
console.log(s.lastIndexOf("b"));
console.log(s.indexOf("x"));
"#), vec!["1", "4", "-1"]);
}

#[test]
fn indexof_with_start_position() {
    assert_eq!(run_js(r#"
const s = "abcabc";
console.log(s.indexOf("a", 1));
console.log(s.indexOf("a", 4));
"#), vec!["3", "-1"]);
}

// ── string transformation ─────────────────────────────────────────────────────

#[test]
fn touppercase_tolowercase() {
    assert_eq!(run_js(r#"
console.log("Hello World".toUpperCase());
console.log("Hello World".toLowerCase());
"#), vec!["HELLO WORLD", "hello world"]);
}

#[test]
fn trim_removes_whitespace_both_sides() {
    assert_eq!(run_js(r#"
console.log("  hello  ".trim());
console.log("\t\nhello\n\t".trim());
"#), vec!["hello", "hello"]);
}

// ── template expressions ──────────────────────────────────────────────────────

#[test]
fn template_with_arithmetic() {
    assert_eq!(run_js(r#"
const a = 3, b = 4;
console.log(`hypotenuse: ${Math.sqrt(a**2 + b**2)}`);
"#), vec!["hypotenuse: 5"]);
}

#[test]
fn template_with_array_method() {
    assert_eq!(run_js(r#"
const nums = [1, 2, 3, 4, 5];
console.log(`sum: ${nums.reduce((a, b) => a + b, 0)}`);
"#), vec!["sum: 15"]);
}

// ── string comparison ─────────────────────────────────────────────────────────

#[test]
fn string_strict_equality() {
    assert_eq!(run_js(r#"
console.log("abc" === "abc");
console.log("abc" === "ABC");
console.log(new String("abc") === "abc");
"#), vec!["true", "false", "false"]);
}

// ── charCodeAt and fromCharCode ───────────────────────────────────────────────

#[test]
fn charcodeat_basic() {
    assert_eq!(run_js(r#"
console.log("A".charCodeAt(0));
console.log("a".charCodeAt(0));
console.log("Z".charCodeAt(0));
"#), vec!["65", "97", "90"]);
}

#[test]
fn from_charcode_builds_string() {
    assert_eq!(run_js(r#"
console.log(String.fromCharCode(72, 101, 108, 108, 111));
"#), vec!["Hello"]);
}

// ── String.prototype.normalize ────────────────────────────────────────────────

#[test]
fn normalize_nfc_composes() {
    assert_eq!(run_js(r#"
const decomposed = "e\u0301"; // e + combining acute
const composed = decomposed.normalize("NFC");
console.log(composed.length);
console.log(composed === "\u00E9"); // é
"#), vec!["1", "true"]);
}

// ── slice and indexOf combination ─────────────────────────────────────────────

#[test]
fn extract_between_delimiters() {
    assert_eq!(run_js(r#"
const s = "prefix[content]suffix";
const start = s.indexOf("[") + 1;
const end = s.indexOf("]");
console.log(s.slice(start, end));
"#), vec!["content"]);
}
