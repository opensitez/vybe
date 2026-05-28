/// RegExp /v flag (Unicode Sets, ES2024) — set notation, string properties,
/// intersection, subtraction, nested classes, \p{} in v-mode, case-insensitive with v.

use super::helpers::run_js;

// ── basic /v flag ─────────────────────────────────────────────────────────────

#[test]
fn v_flag_creates_valid_regex() {
    assert_eq!(run_js(r#"
const re = /[abc]/v;
console.log(re.flags.includes("v"));
console.log(re.test("a"));
console.log(re.test("d"));
"#), vec!["true", "true", "false"]);
}

#[test]
fn v_flag_matches_unicode_property() {
    assert_eq!(run_js(r#"
const re = /^\p{Letter}+$/v;
console.log(re.test("hello"));
console.log(re.test("abc123"));
"#), vec!["true", "false"]);
}

#[test]
fn v_flag_matches_emoji_via_property() {
    assert_eq!(run_js(r#"
const re = /^\p{Emoji}+$/v;
console.log(re.test("😀😁"));
console.log(re.test("abc"));
"#), vec!["true", "false"]);
}

// ── set intersection ──────────────────────────────────────────────────────────

#[test]
fn v_flag_set_intersection_ascii_letter() {
    assert_eq!(run_js(r#"
// ASCII letters = intersection of Letter and ASCII
const re = /^[\p{Letter}&&\p{ASCII}]+$/v;
console.log(re.test("hello"));
console.log(re.test("café"));
"#), vec!["true", "false"]);
}

#[test]
fn v_flag_set_intersection_digit_and_range() {
    assert_eq!(run_js(r#"
// digits 5-9 via intersection
const re = /^[\p{Decimal_Number}&&[5-9]]+$/v;
console.log(re.test("579"));
console.log(re.test("1234"));
"#), vec!["true", "false"]);
}

// ── set subtraction ───────────────────────────────────────────────────────────

#[test]
fn v_flag_set_subtraction_removes_chars() {
    assert_eq!(run_js(r#"
// lowercase letters minus vowels
const re = /^[a-z--[aeiou]]+$/v;
console.log(re.test("bcdf"));
console.log(re.test("bcda"));
"#), vec!["true", "false"]);
}

#[test]
fn v_flag_subtraction_digits_minus_zero() {
    assert_eq!(run_js(r#"
const re = /^[\p{Decimal_Number}--[0]]+$/v;
console.log(re.test("123456789"));
console.log(re.test("1230"));
"#), vec!["true", "false"]);
}

// ── nested character classes ──────────────────────────────────────────────────

#[test]
fn v_flag_nested_class_union() {
    assert_eq!(run_js(r#"
// digits or hex letters
const re = /^[[0-9][a-fA-F]]+$/v;
console.log(re.test("deadbeef123"));
console.log(re.test("xyz"));
"#), vec!["true", "false"]);
}

// ── case-insensitive with /v ──────────────────────────────────────────────────

#[test]
fn v_flag_with_case_insensitive() {
    assert_eq!(run_js(r#"
const re = /^[a-z]+$/vi;
console.log(re.test("HELLO"));
console.log(re.test("Hello"));
console.log(re.test("123"));
"#), vec!["true", "true", "false"]);
}

// ── string properties \p{RGI_Emoji} ─────────────────────────────────────────

#[test]
fn v_flag_string_property_rgi_emoji() {
    assert_eq!(run_js(r#"
// RGI_Emoji can match multi-code-point emoji sequences
const re = /^\p{RGI_Emoji}$/v;
console.log(re.test("😀"));
"#), vec!["true"]);
}

// ── match all with /gv ────────────────────────────────────────────────────────

#[test]
fn v_flag_global_match_all() {
    assert_eq!(run_js(r#"
const re = /[\p{Letter}&&\p{ASCII}]+/gv;
const matches = "hello world 123".match(re);
console.log(matches.join(","));
"#), vec!["hello,world"]);
}

// ── /v vs /u compatibility ────────────────────────────────────────────────────

#[test]
fn v_flag_not_combinable_with_u_flag() {
    assert_eq!(run_js(r#"
let threw = false;
try { new RegExp(".", "uv"); } catch (e) { threw = true; }
console.log(threw);
"#), vec!["true"]);
}

#[test]
fn v_flag_allows_unescaped_dollar_in_class() {
    assert_eq!(run_js(r#"
// In /v, $ doesn't need escaping inside character classes
const re = /^[$€£]+$/v;
console.log(re.test("$€£"));
console.log(re.test("abc"));
"#), vec!["true", "false"]);
}

// ── complement of property ────────────────────────────────────────────────────

#[test]
fn v_flag_negated_property_escape() {
    assert_eq!(run_js(r#"
const re = /^\P{Letter}+$/v;
console.log(re.test("123 !@#"));
console.log(re.test("abc"));
"#), vec!["true", "false"]);
}
