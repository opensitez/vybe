/// Regex named capture groups, lookbehind assertions, /d flag (match indices),
/// named backreferences, replace with named groups, RegExp.exec loop patterns.
use super::helpers::run_js;

// ── named capture groups ──────────────────────────────────────────────────────

#[test]
fn named_group_basic_match() {
    assert_eq!(
        run_js(
            r#"
const re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/;
const m = re.exec("2024-01-15");
console.log(m.groups.year);
console.log(m.groups.month);
console.log(m.groups.day);
"#
        ),
        vec!["2024", "01", "15"]
    );
}

#[test]
fn named_group_via_destructuring() {
    assert_eq!(
        run_js(
            r#"
const { groups: { first, last } } = /(?<first>\w+) (?<last>\w+)/.exec("John Doe");
console.log(first);
console.log(last);
"#
        ),
        vec!["John", "Doe"]
    );
}

#[test]
fn named_group_in_string_match() {
    assert_eq!(
        run_js(
            r#"
const m = "2024-03-21".match(/(?<y>\d{4})-(?<m>\d{2})-(?<d>\d{2})/);
console.log(m.groups.y);
"#
        ),
        vec!["2024"]
    );
}

#[test]
fn named_group_undefined_when_not_matched_optional() {
    assert_eq!(
        run_js(
            r#"
const re = /(?<a>\d+)?(?<b>[a-z]+)/;
const m = re.exec("hello");
console.log(m.groups.a);
console.log(m.groups.b);
"#
        ),
        vec!["undefined", "hello"]
    );
}

// ── named backreferences ──────────────────────────────────────────────────────

#[test]
fn named_backreference_in_same_pattern() {
    assert_eq!(
        run_js(
            r#"
const re = /(?<q>["']).*?\k<q>/;
console.log(re.test('"hello"'));
console.log(re.test("'world'"));
console.log(re.test('"mixed\''));
"#
        ),
        vec!["true", "true", "false"]
    );
}

#[test]
fn named_backreference_repeated_word() {
    assert_eq!(
        run_js(
            r#"
const re = /\b(?<word>\w+)\s+\k<word>\b/i;
const m = re.exec("the the river");
console.log(m ? m.groups.word : "no match");
"#
        ),
        vec!["the"]
    );
}

// ── replace with named groups ─────────────────────────────────────────────────

#[test]
fn replace_uses_named_group_in_template() {
    assert_eq!(
        run_js(
            r#"
const re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/;
const result = "2024-01-15".replace(re, "$<day>/$<month>/$<year>");
console.log(result);
"#
        ),
        vec!["15/01/2024"]
    );
}

#[test]
fn replace_function_receives_groups_object() {
    assert_eq!(
        run_js(
            r#"
const result = "John Smith".replace(
    /(?<first>\w+) (?<last>\w+)/,
    (_, first, last, _offset, _str, groups) => groups.last + ", " + groups.first
);
console.log(result);
"#
        ),
        vec!["Smith, John"]
    );
}

// ── lookbehind assertions ─────────────────────────────────────────────────────

#[test]
fn positive_lookbehind_matches_after_prefix() {
    assert_eq!(
        run_js(
            r#"
const re = /(?<=\$)\d+/;
const m = re.exec("$100 and $200");
console.log(m[0]);
"#
        ),
        vec!["100"]
    );
}

#[test]
fn negative_lookbehind_excludes_prefix() {
    assert_eq!(
        run_js(
            r#"
const re = /(?<!\$)\d+/;
const m = "price: 100, discount: $20".match(re);
console.log(m[0]);
"#
        ),
        vec!["100"]
    );
}

#[test]
fn lookbehind_with_global_flag_all_matches() {
    assert_eq!(
        run_js(
            r#"
const matches = [..."$10, $20, €30".matchAll(/(?<=\$)\d+/g)].map(m => m[0]);
console.log(matches.join(","));
"#
        ),
        vec!["10,20"]
    );
}

#[test]
fn lookahead_positive_matches_before_suffix() {
    assert_eq!(
        run_js(
            r#"
const re = /\d+(?=px)/;
const m = re.exec("12px and 30em");
console.log(m[0]);
"#
        ),
        vec!["12"]
    );
}

#[test]
fn lookahead_negative_excludes_suffix() {
    assert_eq!(
        run_js(
            r#"
const matches = [..."12px 30em 5px".matchAll(/\d+(?!px)\b/g)].map(m => m[0]);
console.log(matches.join(","));
"#
        ),
        vec!["30"]
    );
}

// ── /d flag (match indices) ───────────────────────────────────────────────────

#[test]
fn d_flag_provides_indices_array() {
    assert_eq!(
        run_js(
            r#"
const re = /(\w+)/d;
const m = re.exec("hello world");
console.log(m.indices[0][0]);
console.log(m.indices[0][1]);
"#
        ),
        vec!["0", "5"]
    );
}

#[test]
fn d_flag_provides_named_group_indices() {
    assert_eq!(
        run_js(
            r#"
const re = /(?<word>\w+)/d;
const m = re.exec("hello world");
console.log(m.indices.groups.word[0]);
console.log(m.indices.groups.word[1]);
"#
        ),
        vec!["0", "5"]
    );
}

// ── RegExp.exec loop ──────────────────────────────────────────────────────────

#[test]
fn exec_loop_with_global_flag() {
    assert_eq!(
        run_js(
            r#"
const re = /\d+/g;
const text = "a1b22c333";
const results = [];
let m;
while ((m = re.exec(text)) !== null) {
    results.push(m[0]);
}
console.log(results.join(","));
"#
        ),
        vec!["1,22,333"]
    );
}

#[test]
fn exec_loop_with_named_groups() {
    assert_eq!(
        run_js(
            r#"
const re = /(?<n>\d+)/g;
const results = [];
let m;
while ((m = re.exec("a1b22c333")) !== null) {
    results.push(m.groups.n);
}
console.log(results.join(","));
"#
        ),
        vec!["1,22,333"]
    );
}

// ── sticky flag with lastIndex ────────────────────────────────────────────────

#[test]
fn sticky_flag_matches_from_lastindex() {
    assert_eq!(
        run_js(
            r#"
const re = /\d+/y;
re.lastIndex = 2;
const m = re.exec("ab123");
console.log(m[0]);
console.log(re.lastIndex);
"#
        ),
        vec!["123", "5"]
    );
}

#[test]
fn sticky_flag_fails_if_not_at_lastindex() {
    assert_eq!(
        run_js(
            r#"
const re = /\d+/y;
re.lastIndex = 0;
const m = re.exec("ab123");
console.log(m);
"#
        ),
        vec!["null"]
    );
}

// ── matchAll with named groups ────────────────────────────────────────────────

#[test]
fn matchall_with_named_capture_groups() {
    assert_eq!(
        run_js(
            r#"
const re = /(?<key>\w+)=(?<val>\w+)/g;
const results = [];
for (const m of "a=1 b=2 c=3".matchAll(re)) {
    results.push(m.groups.key + ":" + m.groups.val);
}
console.log(results.join(","));
"#
        ),
        vec!["a:1,b:2,c:3"]
    );
}

// ── Unicode category escapes ──────────────────────────────────────────────────

#[test]
fn unicode_property_escape_letter() {
    assert_eq!(
        run_js(
            r#"
const re = /\p{L}+/u;
const m = re.exec("hello123");
console.log(m[0]);
"#
        ),
        vec!["hello"]
    );
}

#[test]
fn unicode_property_escape_digit() {
    assert_eq!(
        run_js(
            r#"
const re = /\p{Nd}+/u;
const m = re.exec("abc456def");
console.log(m[0]);
"#
        ),
        vec!["456"]
    );
}
