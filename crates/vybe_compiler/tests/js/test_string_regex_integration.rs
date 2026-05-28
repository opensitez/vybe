/// String search, replace, split — regex, function replacers, capture groups

use super::helpers::run_js;

#[test]
fn replace_with_function() {
    assert_eq!(run_js(r#"
const result = "hello world".replace(/(\w+)/g, m => m.toUpperCase());
console.log(result);
"#), vec!["HELLO WORLD"]);
}

#[test]
fn replace_function_receives_match_groups_offset() {
    assert_eq!(run_js(r#"
const result = "2024-06-15".replace(
    /(\d{4})-(\d{2})-(\d{2})/,
    (full, year, month, day) => `${day}/${month}/${year}`
);
console.log(result);
"#), vec!["15/06/2024"]);
}

#[test]
fn replace_named_group_in_replacement() {
    assert_eq!(run_js(r#"
const result = "2024-06-15".replace(
    /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/,
    (_, y, m, d) => `${d}/${m}/${y}`
);
console.log(result);
"#), vec!["15/06/2024"]);
}

#[test]
fn replace_all_string() {
    assert_eq!(run_js(r#"
console.log("aabbcc".replaceAll("b", "X"));
"#), vec!["aaXXcc"]);
}

#[test]
fn split_with_regex() {
    assert_eq!(run_js(r#"
const parts = "one1two2three3".split(/\d/);
console.log(parts.join("|"));
"#), vec!["one|two|three|"]);
}

#[test]
fn split_with_capture_group() {
    assert_eq!(run_js(r#"
// Capture groups appear in the split result
const parts = "one-two+three".split(/([-+])/);
console.log(parts.join(","));
"#), vec!["one,-,two,+,three"]);
}

#[test]
fn split_with_limit() {
    assert_eq!(run_js(r#"
const parts = "a,b,c,d".split(",", 2);
console.log(parts.length);
console.log(parts.join("|"));
"#), vec!["2", "a|b"]);
}

#[test]
fn search_returns_index() {
    assert_eq!(run_js(r#"
console.log("hello world".search(/world/));
console.log("hello world".search(/xyz/));
"#), vec!["6", "-1"]);
}

#[test]
fn search_ignores_global_flag() {
    assert_eq!(run_js(r#"
const re = /o/g;
re.lastIndex = 5;
console.log("foobar".search(re)); // always from 0
"#), vec!["1"]);
}

#[test]
fn match_with_global_returns_all() {
    assert_eq!(run_js(r#"
const matches = "test1 test2 test3".match(/test\d/g);
console.log(matches.join(","));
"#), vec!["test1,test2,test3"]);
}

#[test]
fn match_without_global_returns_first_with_groups() {
    assert_eq!(run_js(r#"
const m = "2024-06-15".match(/(\d{4})-(\d{2})-(\d{2})/);
console.log(m[0]);
console.log(m[1]);
console.log(m[2]);
"#), vec!["2024-06-15", "2024", "06"]);
}

#[test]
fn replace_with_dollar_signs() {
    assert_eq!(run_js(r#"
// $1 = first capture group, $& = whole match
const result = "hello world".replace(/(\w+)\s(\w+)/, "$2 $1");
console.log(result);
const withFull = "abc".replace(/b/, "[$&]");
console.log(withFull);
"#), vec!["world hello", "a[b]c"]);
}
