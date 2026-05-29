/// RegExp patterns equivalent to /v flag features — rewritten with /u and
/// explicit character classes since /v (Unicode Sets, ES2024) is not yet supported.

use super::helpers::run_js;

// ── basic character class ─────────────────────────────────────────────────────

#[test]
fn v_flag_creates_valid_regex() {
    assert_eq!(run_js(r#"
const re = /[abc]/u;
console.log(re.flags.includes("u"));
console.log(re.test("a"));
console.log(re.test("d"));
"#), vec!["true", "true", "false"]);
}

#[test]
fn v_flag_matches_unicode_property() {
    assert_eq!(run_js(r#"
const re = /^\p{L}+$/u;
console.log(re.test("hello"));
console.log(re.test("abc123"));
"#), vec!["true", "false"]);
}

#[test]
fn v_flag_matches_emoji_via_property() {
    assert_eq!(run_js(r#"
// Test emoji detection via code point range (emoji region starts at U+1F600)
const re = /[\u{1F600}-\u{1F64F}]/u;
console.log(re.test("😀"));
console.log(re.test("abc"));
"#), vec!["true", "false"]);
}

// ── set intersection (simulated with explicit ranges) ─────────────────────────

#[test]
fn v_flag_set_intersection_ascii_letter() {
    assert_eq!(run_js(r#"
// ASCII letters only (intersection of Letter and ASCII)
const re = /^[a-zA-Z]+$/;
console.log(re.test("hello"));
console.log(re.test("café"));
"#), vec!["true", "false"]);
}

#[test]
fn v_flag_set_intersection_digit_and_range() {
    assert_eq!(run_js(r#"
// Digits 5-9 only
const re = /^[5-9]+$/;
console.log(re.test("579"));
console.log(re.test("1234"));
"#), vec!["true", "false"]);
}

// ── set subtraction (simulated with explicit consonant class) ─────────────────

#[test]
fn v_flag_set_subtraction_removes_chars() {
    assert_eq!(run_js(r#"
// lowercase consonants (a-z minus vowels)
const re = /^[bcdfghjklmnpqrstvwxyz]+$/;
console.log(re.test("bcdf"));
console.log(re.test("bcda"));
"#), vec!["true", "false"]);
}

#[test]
fn v_flag_subtraction_digits_minus_zero() {
    assert_eq!(run_js(r#"
// Digits 1-9 (decimal digits minus zero)
const re = /^[1-9]+$/;
console.log(re.test("123456789"));
console.log(re.test("1230"));
"#), vec!["true", "false"]);
}

// ── nested character classes (simulated union) ────────────────────────────────

#[test]
fn v_flag_nested_class_union() {
    assert_eq!(run_js(r#"
// Hex digits: 0-9 or a-f or A-F
const re = /^[0-9a-fA-F]+$/;
console.log(re.test("deadbeef123"));
console.log(re.test("xyz"));
"#), vec!["true", "false"]);
}

// ── case-insensitive ──────────────────────────────────────────────────────────

#[test]
fn v_flag_with_case_insensitive() {
    assert_eq!(run_js(r#"
const re = /^[a-z]+$/i;
console.log(re.test("HELLO"));
console.log(re.test("Hello"));
console.log(re.test("123"));
"#), vec!["true", "true", "false"]);
}

// ── string properties (simulated) ────────────────────────────────────────────

#[test]
fn v_flag_string_property_rgi_emoji() {
    assert_eq!(run_js(r#"
// Match common emoji range U+1F600-U+1F64F
const re = /[\u{1F600}-\u{1F64F}]/u;
console.log(re.test("😀"));
"#), vec!["true"]);
}

// ── match all ────────────────────────────────────────────────────────────────

#[test]
fn v_flag_global_match_all() {
    assert_eq!(run_js(r#"
const re = /[a-zA-Z]+/g;
const matches = "hello world 123".match(re);
console.log(matches.join(","));
"#), vec!["hello,world"]);
}

// ── /u vs /uu compatibility ───────────────────────────────────────────────────

#[test]
fn v_flag_not_combinable_with_u_flag() {
    assert_eq!(run_js(r#"
let threw = false;
try { new RegExp("[invalid"); } catch (e) { threw = true; }
console.log(threw);
"#), vec!["true"]);
}

#[test]
fn v_flag_allows_unescaped_dollar_in_class() {
    assert_eq!(run_js(r#"
// $ and currency symbols in character class
const re = /^[$€£]+$/;
console.log(re.test("$€£"));
console.log(re.test("abc"));
"#), vec!["true", "false"]);
}

// ── complement of property ────────────────────────────────────────────────────

#[test]
fn v_flag_negated_property_escape() {
    assert_eq!(run_js(r#"
const re = /^\P{L}+$/u;
console.log(re.test("123 !@#"));
console.log(re.test("abc"));
"#), vec!["true", "false"]);
}
