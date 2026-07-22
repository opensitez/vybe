use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: RegExp Named Capture Groups & Match Indices (`/d` flag)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_regexp_named_capture_groups_basic() {
    let src = r#"
const re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/;
const match = re.exec("2026-07-22");
console.log(`${match.groups.year}:${match.groups.month}:${match.groups.day}`);
"#;
    assert_eq!(run_js(src), vec!["2026:07:22"]);
}

#[test]
fn test_js_regexp_named_capture_groups_destructuring() {
    let src = r#"
const re = /(?<title>Mr\.|Mrs\.|Dr\.) (?<name>\w+)/;
const { groups: { title, name } } = re.exec("Dr. Watson");
console.log(`${title} ${name}`);
"#;
    assert_eq!(run_js(src), vec!["Dr. Watson"]);
}

#[test]
fn test_js_regexp_named_capture_groups_backreference() {
    let src = r#"
const re = /<(?<tag>\w+)>.*<\/k\k<tag>>/; // \k<tag> backreference
console.log(re.test("<div>Hello</div>") + "|" + re.test("<div>World</span>"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_regexp_match_indices_d_flag() {
    let src = r#"
const re = /b(a)(r)/d;
const match = re.exec("foobar");
console.log(match.indices[0].join(":") + "|" + match.indices[1].join(":"));
"#;
    assert_eq!(run_js(src), vec!["3:6|4:5"]);
}

#[test]
fn test_js_regexp_match_indices_named_groups() {
    let src = r#"
const re = /(?<word>\w+)/d;
const match = re.exec("hello");
console.log(match.indices.groups.word.join(":"));
"#;
    assert_eq!(run_js(src), vec!["0:5"]);
}

#[test]
fn test_js_regexp_named_capture_groups_unmatched_returns_undefined() {
    let src = r#"
const re = /(?<a>x)|(?<b>y)/;
const match = re.exec("x");
console.log(match.groups.a + "|" + match.groups.b);
"#;
    assert_eq!(run_js(src), vec!["x|undefined"]);
}

#[test]
fn test_js_regexp_named_capture_groups_string_replace() {
    let src = r#"
const re = /(?<first>\w+)\s+(?<last>\w+)/;
const res = "John Doe".replace(re, "$<last>, $<first>");
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["Doe, John"]);
}

#[test]
fn test_js_regexp_has_indices_flag_property() {
    let src = r#"
const re = /abc/d;
console.log(re.hasIndices);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_duplicate_named_capture_groups_in_disjunction_es2023() {
    let src = r#"
const re = /(?<date>\d{4}-\d{2})|(?<date>\d{2}\/\d{2})/; // Duplicate name across alternate branches
const m1 = re.exec("2026-07");
const m2 = re.exec("07/22");
console.log(m1.groups.date + "|" + m2.groups.date);
"#;
    assert_eq!(run_js(src), vec!["2026-07|07/22"]);
}

#[test]
fn test_js_regexp_named_capture_groups_null_prototype() {
    let src = r#"
const re = /(?<val>\d+)/;
const match = re.exec("100");
console.log(Object.getPrototypeOf(match.groups) === null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_match_indices_unmatched_group_indices_undefined() {
    let src = r#"
const re = /(?<x>a)|(?<y>b)/d;
const match = re.exec("a");
console.log(match.indices[1].join(":") + "|" + (match.indices[2] === undefined));
"#;
    assert_eq!(run_js(src), vec!["0:1|true"]);
}

#[test]
fn test_js_regexp_string_replace_func_named_groups_param() {
    let src = r#"
const re = /(?<num>\d+)/;
const res = "Item 42".replace(re, (match, p1, offset, string, groups) => {
    return String(Number(groups.num) * 2);
});
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["Item 84"]);
}

#[test]
fn test_js_regexp_named_capture_groups_prototype_flags() {
    let src = r#"
const re = new RegExp("a", "dg");
console.log(re.flags);
"#;
    assert_eq!(run_js(src), vec!["dg"]);
}

#[test]
fn test_js_regexp_exec_null_when_no_match() {
    let src = r#"
const re = /(?<tag>\d+)/;
console.log(re.exec("abc") === null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_named_capture_groups_in_lookahead() {
    let src = r#"
const re = /(?=(?<p>\d+))\d{3}/;
const match = re.exec("12345");
console.log(match.groups.p);
"#;
    assert_eq!(run_js(src), vec!["12345"]);
}

#[test]
fn test_js_regexp_indices_property_has_own_property() {
    let src = r#"
const re = /a/d;
const match = re.exec("a");
console.log(Object.hasOwn(match, "indices"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_match_all_named_groups() {
    let src = r#"
const re = /(?<letter>[a-z])(?<digit>\d)/g;
const matches = [... "a1b2c3".matchAll(re)];
console.log(matches.map(m => `${m.groups.letter}:${m.groups.digit}`).join(","));
"#;
    assert_eq!(run_js(src), vec!["a:1,b:2,c:3"]);
}

#[test]
fn test_js_regexp_invalid_named_backreference_throws_syntaxerror() {
    let src = r#"
try {
    eval("const re = /\\k<missing>/;");
} catch (e) {
    console.log("Invalid Backreference SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Backreference SyntaxError"]);
}

#[test]
fn test_js_regexp_named_group_starts_with_digit_throws_syntaxerror() {
    let src = r#"
try {
    eval("const re = /(?<1group>a)/;");
} catch (e) {
    console.log("Invalid Group Identifier SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Group Identifier SyntaxError"]);
}

#[test]
fn test_js_regexp_unicode_identifier_named_groups() {
    let src = r#"
const re = /(?<π>\d+)/u;
const match = re.exec("314");
console.log(match.groups.π);
"#;
    assert_eq!(run_js(src), vec!["314"]);
}
