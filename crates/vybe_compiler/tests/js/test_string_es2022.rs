use super::helpers::run_js;

// ── String.prototype.at() ─────────────────────────────────
#[test]
fn string_at_positive_index() {
    assert_eq!(
        run_js(
            r#"
console.log("hello".at(1));
"#
        ),
        vec!["e"]
    );
}

#[test]
fn string_at_negative_index() {
    assert_eq!(
        run_js(
            r#"
console.log("hello".at(-1));
console.log("hello".at(-2));
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
console.log("hi".at(100) === undefined);
"#
        ),
        vec!["true"]
    );
}

// ── String.prototype.replaceAll ───────────────────────────
#[test]
fn string_replaceall_basic() {
    assert_eq!(
        run_js(
            r#"
console.log("a-b-c-d".replaceAll("-", "_"));
"#
        ),
        vec!["a_b_c_d"]
    );
}

#[test]
fn string_replaceall_with_empty() {
    assert_eq!(
        run_js(
            r#"
console.log("hello".replaceAll("l", ""));
"#
        ),
        vec!["heo"]
    );
}

#[test]
fn string_replaceall_with_function() {
    assert_eq!(
        run_js(
            r#"
const result = "one two one".replaceAll("one", m => m.toUpperCase());
console.log(result);
"#
        ),
        vec!["ONE two ONE"]
    );
}

// ── String.prototype.matchAll ─────────────────────────────
#[test]
fn string_matchall_returns_all_matches() {
    assert_eq!(
        run_js(
            r#"
const str = "cat bat sat";
const matches = [...str.matchAll(/[a-z]at/g)];
console.log(matches.length);
console.log(matches[0][0]);
"#
        ),
        vec!["3", "cat"]
    );
}

#[test]
fn string_matchall_with_capture_groups() {
    assert_eq!(
        run_js(
            r#"
const str = "2024-01-15 2024-12-31";
const matches = [...str.matchAll(/(\d{4})-(\d{2})-(\d{2})/g)];
console.log(matches.length);
console.log(matches[0][1]);
"#
        ),
        vec!["2", "2024"]
    );
}

#[test]
fn string_matchall_indices() {
    assert_eq!(
        run_js(
            r#"
const matches = [...("abcabc".matchAll(/a/g))];
const indices = matches.map(m => m.index).join(",");
console.log(indices);
"#
        ),
        vec!["0,3"]
    );
}

// ── String.raw ────────────────────────────────────────────
#[test]
fn string_raw_preserves_backslashes() {
    assert_eq!(
        run_js(
            r#"
const path = String.raw`C:\Users\name\file.txt`;
console.log(path);
"#
        ),
        vec!["C:\\Users\\name\\file.txt"]
    );
}

#[test]
fn string_raw_no_newline_escapes() {
    assert_eq!(
        run_js(
            r#"
const s = String.raw`line1\nline2`;
console.log(s.includes("\\n"));
"#
        ),
        vec!["true"]
    );
}

// ── String.prototype.trimStart/trimEnd ────────────────────
#[test]
fn string_trimstart_removes_leading() {
    assert_eq!(
        run_js(
            r#"
console.log("   hello   ".trimStart());
"#
        ),
        vec!["hello   "]
    );
}

#[test]
fn string_trimend_removes_trailing() {
    assert_eq!(
        run_js(
            r#"
console.log("   hello   ".trimEnd());
"#
        ),
        vec!["   hello"]
    );
}

#[test]
fn string_trimstart_trimend_aliases() {
    assert_eq!(
        run_js(
            r#"
const s = "  test  ";
console.log(s.trimStart() === s.trimLeft());
console.log(s.trimEnd() === s.trimRight());
"#
        ),
        vec!["true", "true"]
    );
}

// ── String.prototype.padStart/padEnd ──────────────────────
#[test]
fn string_padstart_pads_to_length() {
    assert_eq!(
        run_js(
            r#"
console.log("5".padStart(3, "0"));
"#
        ),
        vec!["005"]
    );
}

#[test]
fn string_padend_pads_to_length() {
    assert_eq!(
        run_js(
            r#"
console.log("hi".padEnd(5, "."));
"#
        ),
        vec!["hi..."]
    );
}

#[test]
fn string_padstart_already_long_enough() {
    assert_eq!(
        run_js(
            r#"
console.log("hello".padStart(3, "x"));
"#
        ),
        vec!["hello"]
    );
}

// ── String.prototype.repeat ───────────────────────────────
#[test]
fn string_repeat_count() {
    assert_eq!(
        run_js(
            r#"
console.log("abc".repeat(3));
"#
        ),
        vec!["abcabcabc"]
    );
}

#[test]
fn string_repeat_zero() {
    assert_eq!(
        run_js(
            r#"
console.log("abc".repeat(0));
"#
        ),
        vec![""]
    );
}

// ── String.prototype.startsWith/endsWith ─────────────────
#[test]
fn string_startswith_with_position() {
    assert_eq!(
        run_js(
            r#"
console.log("hello world".startsWith("world", 6));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn string_endswith_with_length() {
    assert_eq!(
        run_js(
            r#"
console.log("hello world".endsWith("hello", 5));
"#
        ),
        vec!["true"]
    );
}

// ── Template literals ─────────────────────────────────────
#[test]
fn tagged_template_basic() {
    assert_eq!(
        run_js(
            r#"
function tag(strings, ...values) {
  return strings.raw[0] + values[0];
}
const name = "World";
console.log(tag`Hello ${name}`);
"#
        ),
        vec!["Hello World"]
    );
}

#[test]
fn tagged_template_highlight() {
    assert_eq!(
        run_js(
            r#"
function highlight(strings, ...vals) {
  return strings.reduce((acc, str, i) => acc + str + (vals[i] !== undefined ? "[" + vals[i] + "]" : ""), "");
}
const a = 1, b = 2;
console.log(highlight`sum of ${a} and ${b} is ${a + b}`);
"#
        ),
        vec!["sum of [1] and [2] is [3]"]
    );
}

// ── String.prototype.normalize ────────────────────────────
#[test]
fn string_normalize_nfc() {
    assert_eq!(
        run_js(
            r#"
const s = "é";
console.log(s.normalize("NFC") === "é");
"#
        ),
        vec!["true"]
    );
}

// ── String Unicode ────────────────────────────────────────
#[test]
fn string_codepointat_emoji() {
    assert_eq!(
        run_js(
            r#"
const cp = "A".codePointAt(0);
console.log(cp);
"#
        ),
        vec!["65"]
    );
}

#[test]
fn string_fromcodepoint_basic() {
    assert_eq!(
        run_js(
            r#"
console.log(String.fromCodePoint(65, 66, 67));
"#
        ),
        vec!["ABC"]
    );
}

#[test]
fn string_includes_case_sensitive() {
    assert_eq!(
        run_js(
            r#"
console.log("Hello World".includes("World"));
console.log("Hello World".includes("world"));
"#
        ),
        vec!["true", "false"]
    );
}
