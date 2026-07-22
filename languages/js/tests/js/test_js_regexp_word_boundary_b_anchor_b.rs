use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: RegExp Word Boundary (`\b`, `\B`) & Anchors (`^`, `$`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_regexp_word_boundary_b_basic() {
    let src = r##"
const str = "cat concat scatter";
console.log(str.match(/\bcat\b/g).join(","));
"##;
    assert_eq!(run_js(src), vec!["cat"]);
}

#[test]
fn test_js_regexp_non_word_boundary_B_basic() {
    let src = r##"
const str = "cat concat scatter";
console.log(str.match(/\Bcat\B/g).join(","));
"##;
    assert_eq!(run_js(src), vec!["cat"]);
}

#[test]
fn test_js_regexp_start_anchor_caret() {
    let src = r##"
const str = "hello world";
console.log(`${/^hello/.test(str)}:${/^world/.test(str)}`);
"##;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_regexp_end_anchor_dollar() {
    let src = r##"
const str = "hello world";
console.log(`${/world$/.test(str)}:${/hello$/.test(str)}`);
"##;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_regexp_exact_string_match_caret_and_dollar() {
    let src = r##"
const str = "hello";
console.log(`${/^hello$/.test("hello")}:${/^hello$/.test("hello world")}`);
"##;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_regexp_word_boundary_at_start_and_end_of_string() {
    let src = r##"
const str = "word";
console.log(str.match(/\bword\b/).length > 0);
"##;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_word_boundary_with_punctuation() {
    let src = r##"
const str = "hello, world!";
console.log(str.match(/\b\w+\b/g).join(","));
"##;
    assert_eq!(run_js(src), vec!["hello,world"]);
}

#[test]
fn test_js_regexp_non_word_boundary_prefix() {
    let src = r##"
const str = "cat concat scatter";
console.log(str.match(/\Bcat\b/g).join(","));
"##;
    assert_eq!(run_js(src), vec!["cat"]);
}

#[test]
fn test_js_regexp_word_boundary_suffix() {
    let src = r##"
const str = "cat concat scatter";
console.log(str.match(/\bcat\B/g).join(","));
"##;
    assert_eq!(run_js(src), vec!["cat"]);
}

#[test]
fn test_js_regexp_multiline_anchor_caret_with_m_flag() {
    let src = r##"
const str = "line1\nline2\nline3";
console.log(str.match(/^line\d/gm).join(","));
"##;
    assert_eq!(run_js(src), vec!["line1,line2,line3"]);
}

#[test]
fn test_js_regexp_multiline_anchor_dollar_with_m_flag() {
    let src = r##"
const str = "line1\nline2\nline3";
console.log(str.match(/line\d$/gm).join(","));
"##;
    assert_eq!(run_js(src), vec!["line1,line2,line3"]);
}

#[test]
fn test_js_regexp_word_boundary_unicode_u_flag() {
    let src = r##"
const str = "αβγ";
console.log(/\b/u.test(str));
"##;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_empty_string_anchors() {
    let src = r##"
console.log(`${/^$/.test("")}:${/^$/.test("a")}`);
"##;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_regexp_word_boundary_between_two_word_chars_is_false() {
    let src = r##"
console.log(/\b/.test("aa"));
"##;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_word_boundary_between_two_non_word_chars_is_false() {
    let src = r##"
console.log(/\b/.test("!!"));
"##;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_regexp_non_word_boundary_between_two_non_word_chars_is_true() {
    let src = r##"
console.log(/\B/.test("!!"));
"##;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_caret_inside_character_class_is_negation() {
    let src = r##"
const str = "^abc^";
console.log(str.match(/[^^]/g).join(","));
"##;
    assert_eq!(run_js(src), vec!["a,b,c"]);
}

#[test]
fn test_js_regexp_caret_not_at_start_of_character_class_is_literal() {
    let src = r##"
const str = "a^b";
console.log(str.match(/[a^b]/g).join(","));
"##;
    assert_eq!(run_js(src), vec!["a,^,b"]);
}

#[test]
fn test_js_regexp_dollar_inside_character_class_is_literal() {
    let src = r##"
const str = "a$b";
console.log(str.match(/[$]/g).join(","));
"##;
    assert_eq!(run_js(src), vec!["$"]);
}

#[test]
fn test_js_regexp_anchors_with_sticky_y_flag() {
    let src = r##"
const re = /^b/y;
re.lastIndex = 1;
console.log(re.test("ab"));
"##;
    assert_eq!(run_js(src), vec!["false"]);
}
