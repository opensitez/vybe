/// Regex string methods — match, matchAll, replace with function, search,
/// split with regex, regex exec, lastIndex behavior, dotAll flag, unicode flag.
use super::helpers::run_js;

// ── String.match ──────────────────────────────────────────────────────────────

#[test]
fn match_without_global_returns_first_with_groups() {
    assert_eq!(
        run_js(
            r#"
const result = "hello world".match(/(\w+)\s(\w+)/);
console.log(result[0]);
console.log(result[1]);
console.log(result[2]);
console.log(result.index);
"#
        ),
        vec!["hello world", "hello", "world", "0"]
    );
}

#[test]
fn match_with_global_returns_all_matches() {
    assert_eq!(
        run_js(
            r#"
const result = "cat bat sat".match(/[a-z]at/g);
console.log(result.join(","));
"#
        ),
        vec!["cat,bat,sat"]
    );
}

#[test]
fn match_returns_null_when_no_match() {
    assert_eq!(
        run_js(
            r#"
const result = "hello".match(/xyz/);
console.log(result);
"#
        ),
        vec!["null"]
    );
}

// ── String.search ─────────────────────────────────────────────────────────────

#[test]
fn search_returns_index_of_first_match() {
    assert_eq!(
        run_js(
            r#"
console.log("hello world".search(/world/));
console.log("hello world".search(/xyz/));
"#
        ),
        vec!["6", "-1"]
    );
}

#[test]
fn search_ignores_global_flag() {
    assert_eq!(
        run_js(
            r#"
// search always returns first match index, global flag doesn't matter
const re = /\d+/g;
re.lastIndex = 5; // should be ignored
console.log("abc123def456".search(re));
"#
        ),
        vec!["3"]
    );
}

// ── String.replace with regex ─────────────────────────────────────────────────

#[test]
fn replace_first_match_without_global() {
    assert_eq!(
        run_js(
            r#"
const result = "aababc".replace(/a/, "X");
console.log(result);
"#
        ),
        vec!["Xababc"]
    );
}

#[test]
fn replace_all_matches_with_global() {
    assert_eq!(
        run_js(
            r#"
const result = "aababc".replace(/a/g, "X");
console.log(result);
"#
        ),
        vec!["XXbXbc"]
    );
}

#[test]
fn replace_with_capture_group_reference() {
    assert_eq!(
        run_js(
            r#"
const result = "2024-01-15".replace(/(\d{4})-(\d{2})-(\d{2})/, "$3/$2/$1");
console.log(result);
"#
        ),
        vec!["15/01/2024"]
    );
}

#[test]
fn replace_function_receives_match_groups_index() {
    assert_eq!(
        run_js(
            r#"
const result = "hello world".replace(/(\w+)/g, (match, group1, index) => {
    return `[${match}@${index}]`;
});
console.log(result);
"#
        ),
        vec!["[hello@0] [world@6]"]
    );
}

// ── String.split with regex ───────────────────────────────────────────────────

#[test]
fn split_by_regex_multiple_separators() {
    assert_eq!(
        run_js(
            r#"
const parts = "one1two2three3".split(/\d/);
console.log(parts.join(","));
"#
        ),
        vec!["one,two,three,"]
    );
}

#[test]
fn split_regex_with_capture_group() {
    assert_eq!(
        run_js(
            r#"
// Capture groups are included in result
const parts = "a1b2c".split(/(\d)/);
console.log(parts.join(","));
"#
        ),
        vec!["a,1,b,2,c"]
    );
}

// ── RegExp.exec ───────────────────────────────────────────────────────────────

#[test]
fn exec_returns_match_object() {
    assert_eq!(
        run_js(
            r#"
const re = /(\d+)/;
const result = re.exec("abc123def");
console.log(result[0]);
console.log(result[1]);
console.log(result.index);
"#
        ),
        vec!["123", "123", "3"]
    );
}

#[test]
fn exec_with_global_advances_lastindex() {
    assert_eq!(
        run_js(
            r#"
const re = /\d+/g;
const str = "1 2 3";
const matches = [];
let m;
while ((m = re.exec(str)) !== null) {
    matches.push(m[0]);
}
console.log(matches.join(","));
"#
        ),
        vec!["1,2,3"]
    );
}

// ── dotAll flag ───────────────────────────────────────────────────────────────

#[test]
fn dot_does_not_match_newline_by_default() {
    assert_eq!(
        run_js(
            r#"
const re = /a.b/;
console.log(re.test("a\nb"));
console.log(re.test("acb"));
"#
        ),
        vec!["false", "true"]
    );
}

#[test]
fn dot_all_flag_matches_newline() {
    assert_eq!(
        run_js(
            r#"
const re = /a.b/s;
console.log(re.test("a\nb"));
console.log(re.flags.includes("s"));
"#
        ),
        vec!["true", "true"]
    );
}

// ── Unicode flag ──────────────────────────────────────────────────────────────

#[test]
fn unicode_flag_handles_surrogates() {
    assert_eq!(
        run_js(
            r#"
const re = /./u;
const emoji = "😀";
const match = re.exec(emoji);
// With /u, . matches the whole code point (surrogate pair)
console.log(match[0].length);
"#
        ),
        vec!["2"]
    );
}

#[test]
fn unicode_property_escape_basic() {
    assert_eq!(
        run_js(
            r#"
const re = /\p{Lu}/u; // Uppercase letter
console.log(re.test("A"));
console.log(re.test("a"));
"#
        ),
        vec!["true", "false"]
    );
}

// ── Regex flags combinations ──────────────────────────────────────────────────

#[test]
fn regex_gi_flags_global_insensitive() {
    assert_eq!(
        run_js(
            r#"
const matches = "Hello HELLO hello".match(/hello/gi);
console.log(matches.length);
console.log(matches.join(","));
"#
        ),
        vec!["3", "Hello,HELLO,hello"]
    );
}

#[test]
fn regex_m_flag_multiline() {
    assert_eq!(
        run_js(
            r#"
const re = /^\d+/mg;
const text = "1 hello\n2 world\n3 foo";
const matches = text.match(re);
console.log(matches.join(","));
"#
        ),
        vec!["1,2,3"]
    );
}

// ── Named replacement ─────────────────────────────────────────────────────────

#[test]
fn replace_with_named_capture_in_template() {
    assert_eq!(
        run_js(
            r#"
const result = "2024-01-15".replace(
    /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/,
    "$<month>/$<day>/$<year>"
);
console.log(result);
"#
        ),
        vec!["01/15/2024"]
    );
}
