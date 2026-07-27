/// String methods not heavily covered — at, replaceAll, matchAll,
/// trimStart/trimEnd, padStart/padEnd edge cases, repeat, includes/startsWith/endsWith
/// with position args, slice vs substring vs substr, comparison.
use super::helpers::run_js;

// ── String.prototype.at ───────────────────────────────────────────────────────

#[test]
fn string_at_positive() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
console.log(s.at(0));
console.log(s.at(4));
"#
        ),
        vec!["h", "o"]
    );
}

#[test]
fn string_at_negative() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
console.log(s.at(-1));
console.log(s.at(-3));
"#
        ),
        vec!["o", "l"]
    );
}

#[test]
fn string_at_out_of_bounds() {
    assert_eq!(
        run_js(
            r#"
const s = "hi";
console.log(s.at(10));
console.log(s.at(-10));
"#
        ),
        vec!["undefined", "undefined"]
    );
}

// ── replaceAll ────────────────────────────────────────────────────────────────

#[test]
fn replaceall_replaces_every_occurrence() {
    assert_eq!(
        run_js(
            r#"
const s = "a-b-c-d";
console.log(s.replaceAll("-", "_"));
"#
        ),
        vec!["a_b_c_d"]
    );
}

#[test]
fn replaceall_with_function() {
    assert_eq!(
        run_js(
            r#"
const s = "aabbcc";
console.log(s.replaceAll("b", (m) => m.toUpperCase()));
"#
        ),
        vec!["aaBBcc"]
    );
}

#[test]
fn replaceall_no_matches_unchanged() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
console.log(s.replaceAll("x", "y"));
"#
        ),
        vec!["hello"]
    );
}

// ── matchAll ──────────────────────────────────────────────────────────────────

#[test]
fn matchall_returns_all_matches() {
    assert_eq!(
        run_js(
            r#"
const str = "test1 test2 test3";
const matches = [...str.matchAll(/test(\d)/g)];
console.log(matches.length);
console.log(matches[0][1]);
console.log(matches[2][1]);
"#
        ),
        vec!["3", "1", "3"]
    );
}

#[test]
fn matchall_includes_index() {
    assert_eq!(
        run_js(
            r#"
const str = "aXbXc";
const matches = [...str.matchAll(/X/g)];
console.log(matches[0].index);
console.log(matches[1].index);
"#
        ),
        vec!["1", "3"]
    );
}

// ── trimStart / trimEnd ───────────────────────────────────────────────────────

#[test]
fn trimstart_removes_leading_whitespace() {
    assert_eq!(
        run_js(
            r#"
const s = "   hello   ";
console.log(s.trimStart());
"#
        ),
        vec!["hello   "]
    );
}

#[test]
fn trimend_removes_trailing_whitespace() {
    assert_eq!(
        run_js(
            r#"
const s = "   hello   ";
console.log(s.trimEnd());
"#
        ),
        vec!["   hello"]
    );
}

// ── padStart / padEnd edge cases ──────────────────────────────────────────────

#[test]
fn padstart_with_multi_char_fill() {
    assert_eq!(
        run_js(
            r#"
console.log("5".padStart(5, "0"));
console.log("abc".padStart(7, "xy"));
"#
        ),
        vec!["00005", "xyxyabc"]
    );
}

#[test]
fn padend_with_fill_string() {
    assert_eq!(
        run_js(
            r#"
console.log("1".padEnd(5, "23"));
"#
        ),
        vec!["12323"]
    );
}

#[test]
fn pad_does_not_shorten_longer_strings() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
console.log(s.padStart(3));
console.log(s.padEnd(3));
"#
        ),
        vec!["hello", "hello"]
    );
}

// ── includes/startsWith/endsWith with position ────────────────────────────────

#[test]
fn includes_with_start_position() {
    assert_eq!(
        run_js(
            r#"
const s = "hello world";
console.log(s.includes("hello", 0));
console.log(s.includes("hello", 1));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn startswith_with_position() {
    assert_eq!(
        run_js(
            r#"
const s = "hello world";
console.log(s.startsWith("world", 6));
console.log(s.startsWith("world", 0));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn endswith_with_end_position() {
    assert_eq!(
        run_js(
            r#"
const s = "hello world";
console.log(s.endsWith("hello", 5));
console.log(s.endsWith("world"));
"#
        ),
        vec!["true", "true"]
    );
}

// ── slice edge cases ──────────────────────────────────────────────────────────

#[test]
fn slice_negative_indices() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
console.log(s.slice(-3));
console.log(s.slice(-4, -1));
"#
        ),
        vec!["llo", "ell"]
    );
}

#[test]
fn slice_beyond_length() {
    assert_eq!(
        run_js(
            r#"
const s = "hi";
console.log(s.slice(0, 100));
console.log(s.slice(5));
"#
        ),
        vec!["hi", ""]
    );
}

// ── substring vs slice ────────────────────────────────────────────────────────

#[test]
fn substring_swaps_args_if_start_greater() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
console.log(s.substring(3, 1));
console.log(s.slice(3, 1));
"#
        ),
        vec!["el", ""]
    );
}

#[test]
fn substring_treats_negative_as_zero() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
console.log(s.substring(-1, 3));
"#
        ),
        vec!["hel"]
    );
}

// ── repeat ────────────────────────────────────────────────────────────────────

#[test]
fn repeat_basic() {
    assert_eq!(
        run_js(
            r#"
console.log("ab".repeat(3));
console.log("x".repeat(0));
"#
        ),
        vec!["ababab", ""]
    );
}

// ── String comparison ─────────────────────────────────────────────────────────

#[test]
fn string_comparison_lexicographic() {
    assert_eq!(
        run_js(
            r#"
console.log("apple" < "banana");
console.log("z" > "a");
console.log("abc" === "abc");
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn string_comparison_uppercase_before_lowercase() {
    assert_eq!(
        run_js(
            r#"
// In Unicode, uppercase letters come before lowercase
console.log("A" < "a");
"#
        ),
        vec!["true"]
    );
}

// ── split edge cases ──────────────────────────────────────────────────────────

#[test]
fn split_with_limit() {
    assert_eq!(
        run_js(
            r#"
const parts = "a,b,c,d".split(",", 2);
console.log(parts.join("|"));
console.log(parts.length);
"#
        ),
        vec!["a|b", "2"]
    );
}

#[test]
fn split_empty_string_into_chars() {
    assert_eq!(
        run_js(
            r#"
const chars = "abc".split("");
console.log(chars.join("-"));
"#
        ),
        vec!["a-b-c"]
    );
}

#[test]
fn split_by_regex() {
    assert_eq!(
        run_js(
            r#"
const parts = "one1two2three".split(/\d/);
console.log(parts.join(","));
"#
        ),
        vec!["one,two,three"]
    );
}

#[test]
fn match_without_global_flag() {
    assert_eq!(
        run_js(
            r#"
const match = "abc123def456".match(/\d+/);
console.log(match[0]);
console.log(match.index);
"#
        ),
        vec!["123", "3"]
    );
}

#[test]
fn match_all_digits_with_global_flag() {
    assert_eq!(
        run_js(
            r#"
const matches = "a1 b2 c3".match(/\d/g);
console.log(matches.join("|"));
"#
        ),
        vec!["1|2|3"]
    );
}

#[test]
fn search_index_or_minus_one() {
    assert_eq!(
        run_js(
            r#"
console.log("foobar".search(/bar/));
console.log("foobar".search(/xyz/));
"#
        ),
        vec!["3", "-1"]
    );
}

#[test]
fn replace_replaces_first_match_only() {
    assert_eq!(
        run_js(
            r#"
console.log("a-b-b".replace("-", "_"));
console.log("a-b-b".replace(/-/g, "_"));
"#
        ),
        vec!["a_b-b", "a_b_b"]
    );
}

#[test]
fn substr_and_substring_legacy_compat() {
    assert_eq!(
        run_js(
            r#"
console.log("abcdef".substr(2, 2));
console.log("abcdef".substr(-3));
"#
        ),
        vec!["cd", "def"]
    );
}

#[test]
fn code_point_at_and_from_code_point_boundaries() {
    assert_eq!(
        run_js(
            r#"
console.log("A".codePointAt(0));
const ascii = "ab";
console.log(ascii.codePointAt(0));
console.log(ascii.codePointAt(1));
console.log(String.fromCodePoint(0x10FFFF).length);
"#
        ),
        vec!["65", "97", "98", "2"]
    );
}
