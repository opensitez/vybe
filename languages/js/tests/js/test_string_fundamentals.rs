/// String interpolation, template expressions, multiline, escape sequences,
/// raw strings, unicode escapes, tagged template interactions.
use super::helpers::run_js;

// ── escape sequences in strings ───────────────────────────────────────────────

#[test]
fn escape_sequence_newline_tab() {
    assert_eq!(
        run_js(
            r#"
const s = "line1\nline2\ttab";
const parts = s.split("\n");
console.log(parts[0]);
console.log(parts[1].startsWith("line2\t"));
"#
        ),
        vec!["line1", "true"]
    );
}

#[test]
fn escape_sequence_unicode_code_point() {
    assert_eq!(
        run_js(
            r#"
console.log("\u0041");   // A
console.log("\u{1F600}"); // 😀
console.log("\u{0041}"); // A via curly syntax
"#
        ),
        vec!["A", "😀", "A"]
    );
}

#[test]
fn escape_sequence_hex() {
    assert_eq!(
        run_js(
            r#"
console.log("\x41"); // A
console.log("\x61"); // a
"#
        ),
        vec!["A", "a"]
    );
}

#[test]
fn null_character_escape() {
    assert_eq!(
        run_js(
            r#"
const s = "a\0b";
console.log(s.length);
console.log(s.charCodeAt(1));
"#
        ),
        vec!["3", "0"]
    );
}

// ── string concatenation ──────────────────────────────────────────────────────

#[test]
fn string_concat_method() {
    assert_eq!(
        run_js(
            r#"
const s = "Hello".concat(", ", "World", "!");
console.log(s);
"#
        ),
        vec!["Hello, World!"]
    );
}

#[test]
fn string_plus_coerces_non_strings() {
    assert_eq!(
        run_js(
            r#"
console.log("value: " + 42);
console.log("flag: " + true);
console.log("nothing: " + null);
"#
        ),
        vec!["value: 42", "flag: true", "nothing: null"]
    );
}

// ── string iteration ──────────────────────────────────────────────────────────

#[test]
fn string_for_of_yields_characters() {
    assert_eq!(
        run_js(
            r#"
const chars = [];
for (const c of "abc") chars.push(c);
console.log(chars.join("-"));
"#
        ),
        vec!["a-b-c"]
    );
}

#[test]
fn spread_string_into_array() {
    assert_eq!(
        run_js(
            r#"
const arr = [..."hello"];
console.log(arr.join(","));
"#
        ),
        vec!["h,e,l,l,o"]
    );
}

// ── string searching ──────────────────────────────────────────────────────────

#[test]
fn indexof_and_lastindexof() {
    assert_eq!(
        run_js(
            r#"
const s = "abcabc";
console.log(s.indexOf("b"));
console.log(s.lastIndexOf("b"));
console.log(s.indexOf("x"));
"#
        ),
        vec!["1", "4", "-1"]
    );
}

#[test]
fn indexof_with_start_position() {
    assert_eq!(
        run_js(
            r#"
const s = "abcabc";
console.log(s.indexOf("a", 1));
console.log(s.indexOf("a", 4));
"#
        ),
        vec!["3", "-1"]
    );
}

// ── string transformation ─────────────────────────────────────────────────────

#[test]
fn touppercase_tolowercase() {
    assert_eq!(
        run_js(
            r#"
console.log("Hello World".toUpperCase());
console.log("Hello World".toLowerCase());
"#
        ),
        vec!["HELLO WORLD", "hello world"]
    );
}

#[test]
fn trim_removes_whitespace_both_sides() {
    assert_eq!(
        run_js(
            r#"
console.log("  hello  ".trim());
console.log("\t\nhello\n\t".trim());
"#
        ),
        vec!["hello", "hello"]
    );
}

// ── template expressions ──────────────────────────────────────────────────────

#[test]
fn template_with_arithmetic() {
    assert_eq!(
        run_js(
            r#"
const a = 3, b = 4;
console.log(`hypotenuse: ${Math.sqrt(a**2 + b**2)}`);
"#
        ),
        vec!["hypotenuse: 5"]
    );
}

#[test]
fn template_with_array_method() {
    assert_eq!(
        run_js(
            r#"
const nums = [1, 2, 3, 4, 5];
console.log(`sum: ${nums.reduce((a, b) => a + b, 0)}`);
"#
        ),
        vec!["sum: 15"]
    );
}

// ── string comparison ─────────────────────────────────────────────────────────

#[test]
fn string_strict_equality() {
    assert_eq!(
        run_js(
            r#"
console.log("abc" === "abc");
console.log("abc" === "ABC");
console.log(new String("abc") === "abc");
"#
        ),
        vec!["true", "false", "false"]
    );
}

// ── charCodeAt and fromCharCode ───────────────────────────────────────────────

#[test]
fn charcodeat_basic() {
    assert_eq!(
        run_js(
            r#"
console.log("A".charCodeAt(0));
console.log("a".charCodeAt(0));
console.log("Z".charCodeAt(0));
"#
        ),
        vec!["65", "97", "90"]
    );
}

#[test]
fn from_charcode_builds_string() {
    assert_eq!(
        run_js(
            r#"
console.log(String.fromCharCode(72, 101, 108, 108, 111));
"#
        ),
        vec!["Hello"]
    );
}

// ── String.prototype.normalize ────────────────────────────────────────────────

#[test]
fn normalize_nfc_composes() {
    assert_eq!(
        run_js(
            r#"
const decomposed = "e\u0301"; // e + combining acute
const composed = decomposed.normalize("NFC");
console.log(composed.length);
console.log(composed === "\u00E9"); // é
"#
        ),
        vec!["1", "true"]
    );
}

// ── slice and indexOf combination ─────────────────────────────────────────────

#[test]
fn extract_between_delimiters() {
    assert_eq!(
        run_js(
            r#"
const s = "prefix[content]suffix";
const start = s.indexOf("[") + 1;
const end = s.indexOf("]");
console.log(s.slice(start, end));
"#
        ),
        vec!["content"]
    );
}

#[test]
fn template_literal_with_ternary_and_expression() {
    assert_eq!(
        run_js(
            r#"
const user = { first: "Ada", last: "Lovelace", active: true };
console.log(`${user.first} ${user.last} is ${user.active ? "active" : "inactive"}`);
const total = 3 + 4;
console.log(`total=${total}`);
"#
        ),
        vec!["Ada Lovelace is active", "total=7"]
    );
}

#[test]
fn template_literal_preserves_newlines() {
    assert_eq!(
        run_js(
            r#"
const block = `line1
line2`;
console.log(block.includes("line1"));
console.log(block.includes("line2"));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn string_search_methods_positions() {
    assert_eq!(
        run_js(
            r#"
console.log("javascript".startsWith("java"));
console.log("javascript".startsWith("ava", 1));
console.log("javascript".endsWith("ipt"));
console.log("javascript".endsWith("java", 4));
console.log("javascript".includes("script"));
console.log("javascript".includes("java", 1));
"#
        ),
        vec!["true", "true", "true", "true", "true", "false"]
    );
}

#[test]
fn trim_start_end() {
    assert_eq!(
        run_js(
            r#"
console.log("  spaced  ".trimStart());
console.log("  spaced  ".trimEnd());
const noTrim = "abc".trimStart().trimEnd();
console.log(noTrim);
"#
        ),
        vec!["spaced  ", "  spaced", "abc"]
    );
}

#[test]
fn repeat_negative_throws() {
    assert_eq!(
        run_js(
            r#"
console.log("ha".repeat(3));
try {
    console.log("x".repeat(-1));
} catch (e) {
    console.log(e.name);
}
"#
        ),
        vec!["hahaha", "RangeError"]
    );
}

#[test]
fn string_pad_start_and_pad_end() {
    let src = r#"
console.log("x".padStart(4, "0"));
console.log("x".padEnd(4, "0"));
console.log("x".padStart(2, "ab"));
"#;
    assert_eq!(run_js(src), vec!["000x", "x000", "ax"]);
}

#[test]
fn slice_with_negative_and_bounds() {
    assert_eq!(
        run_js(
            r#"
const s = "abcdef";
console.log(s.slice(-2));
console.log(s.slice(2, 4));
console.log(s.slice(-10, 2));
"#
        ),
        vec!["ef", "cd", "ab"]
    );
}

#[test]
fn string_replace_regex_and_capture_groups() {
    let src = r#"
const input = "a1b2c3";
const replaced = input.replace(/(\d)/g, "[$1]");
console.log(replaced);
console.log("abc123".replace(/(ab)(c)/, "$2-$1"));
"#;
    assert_eq!(run_js(src), vec!["a[1]b[2]c[3]", "c-ab123"]);
}

#[test]
fn string_split_with_limit_and_empty_pattern() {
    let src = r#"
console.log("a,b,c".split(",", 2).join("|"));
console.log("abc".split("").join("-"));
"#;
    assert_eq!(run_js(src), vec!["a|b", "a-b-c"]);
}

#[test]
fn template_raw_keeps_escape_sequences_literal() {
    assert_eq!(
        run_js(
            r#"
const raw = String.raw`line1\nline2`;
console.log(raw.includes("\\n"));
console.log(raw.split("\\n")[0]);
console.log(raw.split("\\n")[1]);
"#
        ),
        vec!["true", "line1", "line2"]
    );
}

#[test]
fn string_index_assignment_is_ignored() {
    assert_eq!(
        run_js(
            r#"
const name = "abc";
console.log(name[0]);
name[0] = "z";
console.log(name);
console.log(name.at(0));
"#
        ),
        vec!["a", "abc", "a"]
    );
}

#[test]
fn code_point_and_from_code_point_behaviors() {
    assert_eq!(
        run_js(
            r#"
console.log(String.fromCodePoint(0x1F600).length);
console.log(String.fromCodePoint(0x1F600).charCodeAt(0));
console.log(String.fromCodePoint(0x1F600).charCodeAt(1));
console.log(String.fromCharCode(0x41).length);
"#,
        ),
        vec!["2", "55357", "56832", "1"]
    );
}
