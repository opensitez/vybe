use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: RegExp Sticky (`/y`) & Global (`/g`) Flag Mechanics (`lastIndex`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_regexp_sticky_flag_strict_position_match() {
    let src = r#"
const re = /foo/y;
re.lastIndex = 3;
console.log(re.test("123foo") + "|lastIndex=" + re.lastIndex);
"#;
    assert_eq!(run_js(src), vec!["true|lastIndex=6"]);
}

#[test]
fn test_js_regexp_sticky_flag_fails_if_not_exact_offset() {
    let src = r#"
const re = /foo/y;
re.lastIndex = 1; // Not matching 'foo' at index 1 -> fails and resets lastIndex to 0!
console.log(re.test("123foo") + "|lastIndex=" + re.lastIndex);
"#;
    assert_eq!(run_js(src), vec!["false|lastIndex=0"]);
}

#[test]
fn test_js_regexp_global_flag_advances_lastindex() {
    let src = r#"
const re = /\w+/g;
const str = "one two";
const m1 = re.exec(str);
const m2 = re.exec(str);
console.log(`${m1[0]}:${re.lastIndex}|${m2[0]}:${re.lastIndex}`);
"#;
    assert_eq!(run_js(src), vec!["one:3|two:7"]);
}

#[test]
fn test_js_regexp_global_flag_resets_lastindex_on_failure() {
    let src = r#"
const re = /\d+/g;
re.test("100"); // lastIndex becomes 3
console.log(re.lastIndex);
re.test("100"); // Search from lastIndex=3 -> fails -> resets to 0!
console.log(re.lastIndex);
"#;
    assert_eq!(run_js(src), vec!["3", "0"]);
}

#[test]
fn test_js_regexp_non_global_non_sticky_ignores_lastindex() {
    let src = r#"
const re = /a/;
re.lastIndex = 2;
const match = re.exec("cat");
console.log(match.index + "|lastIndex=" + re.lastIndex);
"#;
    assert_eq!(run_js(src), vec!["1|lastIndex=2"]);
}

#[test]
fn test_js_regexp_sticky_flag_property() {
    let src = r#"
const re = /abc/y;
console.log(re.sticky + "|" + re.global);
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_regexp_sticky_and_global_combined() {
    let src = r#"
const re = /\d/gy;
const str = "123a45";
const results = [];
let m;
while ((m = re.exec(str)) !== null) {
    results.push(m[0]);
}
console.log(results.join(",") + "|stoppedIndex=" + re.lastIndex);
"#;
    assert_eq!(run_js(src), vec!["1,2,3|stoppedIndex=0"]);
}

#[test]
fn test_js_regexp_sticky_lexer_tokenizer_pattern() {
    let src = r#"
const str = "let x = 10;";
const tokenRe = /\s*([a-z]+|\+|\=|\d+|;)/y;
const tokens = [];

while (tokenRe.lastIndex < str.length) {
    const match = tokenRe.exec(str);
    if (!match) break;
    tokens.push(match[1]);
}
console.log(tokens.join(","));
"#;
    assert_eq!(run_js(src), vec!["let,x,=,10,;"]);
}

#[test]
fn test_js_regexp_lastindex_manual_mutation() {
    let src = r#"
const re = /b/g;
const str = "abcba";
re.lastIndex = 2; // Jump to index 2
const match = re.exec(str);
console.log(match.index); // Finds 'b' at index 3
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_regexp_lastindex_out_of_bounds_resets() {
    let src = r#"
const re = /a/g;
re.lastIndex = 100;
console.log(re.test("abc") + "|lastIndex=" + re.lastIndex);
"#;
    assert_eq!(run_js(src), vec!["false|lastIndex=0"]);
}

#[test]
fn test_js_regexp_sticky_zero_length_match_advances_lastindex() {
    let src = r#"
const re = /^|a/y;
console.log(re.test("a") + "|last=" + re.lastIndex);
console.log(re.test("a") + "|last=" + re.lastIndex);
"#;
    assert_eq!(run_js(src), vec!["true|last=0", "true|last=1"]);
}

#[test]
fn test_js_regexp_global_zero_length_match_advances_lastindex() {
    let src = r#"
const re = /a*/g;
const str = "a";
console.log(`${re.exec(str).index}:${re.lastIndex}`);
console.log(`${re.exec(str).index}:${re.lastIndex}`);
console.log(re.exec(str) === null);
"#;
    assert_eq!(run_js(src), vec!["0:1", "1:1", "true"]);
}

#[test]
fn test_js_regexp_string_match_with_global_returns_array_of_all_matches() {
    let src = r#"
const str = "test1 test2 test3";
const matches = str.match(/test\d/g);
console.log(matches.join(","));
"#;
    assert_eq!(run_js(src), vec!["test1,test2,test3"]);
}

#[test]
fn test_js_regexp_string_match_with_sticky_returns_single_match() {
    let src = r#"
const str = "test1 test2";
const match = str.match(/test\d/y);
console.log(match[0]);
"#;
    assert_eq!(run_js(src), vec!["test1"]);
}

#[test]
fn test_js_regexp_lastindex_non_number_coercion() {
    let src = r#"
const re = /a/g;
re.lastIndex = "1";
console.log(re.exec("cat").index);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_regexp_lastindex_read_only_when_frozen_throws() {
    let src = r#"
const re = /a/g;
Object.freeze(re);
try {
    "use strict";
    re.exec("a");
} catch (e) {
    console.log("Frozen RegExp Exec TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Frozen RegExp Exec TypeError"]);
}

#[test]
fn test_js_regexp_global_flag_string_search_ignores_lastindex() {
    let src = r#"
const re = /b/g;
re.lastIndex = 3;
console.log("abc".search(re)); // search() always starts from index 0 regardless of lastIndex!
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_regexp_sticky_flag_multiline_anchor() {
    let src = r#"
const re = /^bar/ym;
re.lastIndex = 4;
console.log(re.test("foo\nbar"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_constructor_preserves_flags() {
    let src = r#"
const re1 = /abc/gy;
const re2 = new RegExp(re1);
console.log(re2.global + "|" + re2.sticky);
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_regexp_constructor_override_flags() {
    let src = r#"
const re1 = /abc/g;
const re2 = new RegExp(re1, "y");
console.log(re2.global + "|" + re2.sticky);
"#;
    assert_eq!(run_js(src), vec!["false|true"]);
}
