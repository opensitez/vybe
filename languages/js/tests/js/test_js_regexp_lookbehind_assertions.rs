use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: RegExp Lookbehind Assertions (`(?<=...)` & `(?<!...)`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_regexp_positive_lookbehind_basic() {
    let src = r#"
const re = /(?<=\$)\d+/;
const match = re.exec("Price: $100");
console.log(match[0]);
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_regexp_negative_lookbehind_basic() {
    let src = r#"
const re = /(?<!\$)\d+/;
const match = re.exec("Price: 50 USD ($100)");
console.log(match[0]);
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_regexp_lookbehind_variable_length() {
    let src = r#"
const re = /(?<=(a|bb))\d+/; // JS supports variable-length lookbehind!
const m1 = re.exec("a123");
const m2 = re.exec("bb456");
console.log(m1[0] + "|" + m2[0]);
"#;
    assert_eq!(run_js(src), vec!["123|456"]);
}

#[test]
fn test_js_regexp_lookbehind_with_quantifiers() {
    let src = r#"
const re = /(?<=\w+@)\w+\.com/;
const match = re.exec("Contact user@example.com");
console.log(match[0]);
"#;
    assert_eq!(run_js(src), vec!["example.com"]);
}

#[test]
fn test_js_regexp_lookbehind_does_not_consume_characters() {
    let src = r#"
const re = /(?<=#)\w+/;
const match = re.exec("Tag: #coding");
console.log(match[0] + "|index=" + match.index); // Index points to start of matched string 'coding'
"#;
    assert_eq!(run_js(src), vec!["coding|index=6"]);
}

#[test]
fn test_js_regexp_lookbehind_combined_with_lookahead() {
    let src = r#"
const re = /(?<=<tag>).*(?=<\/tag>)/;
const match = re.exec("<tag>Inner Content</tag>");
console.log(match[0]);
"#;
    assert_eq!(run_js(src), vec!["Inner Content"]);
}

#[test]
fn test_js_regexp_lookbehind_capture_groups() {
    let src = r#"
const re = /(?<=(?<prefix>[A-Z]{2}))\d{3}/;
const match = re.exec("ID: AB123");
console.log(match[0] + "|prefix=" + match.groups.prefix);
"#;
    assert_eq!(run_js(src), vec!["123|prefix=AB"]);
}

#[test]
fn test_js_regexp_negative_lookbehind_word_boundary() {
    let src = r#"
const re = /(?<!cat)dog/;
console.log(re.test("hotdog") + "|" + re.test("catdog"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_regexp_lookbehind_at_start_of_string() {
    let src = r#"
const re = /(?<=^)\w+/;
const match = re.exec("Hello World");
console.log(match[0]);
"#;
    assert_eq!(run_js(src), vec!["Hello"]);
}

#[test]
fn test_js_regexp_lookbehind_in_string_replace() {
    let src = r#"
const re = /(?<=\$)(\d+)/g;
const res = "Item1: $10, Item2: $20, Fee: 5".replace(re, (m, val) => String(Number(val) * 2));
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["Item1: $20, Item2: $40, Fee: 5"]);
}

#[test]
fn test_js_regexp_lookbehind_backreference() {
    let src = r#"
const re = /(?<quote>['"])\w+(?<=\k<quote>)/; // Ensures matching quotes
console.log(re.test("'valid'") + "|" + re.test("\"valid\""));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_regexp_lookbehind_non_greedy_quantifier() {
    let src = r#"
const re = /(?<=a+?)\d+/;
const match = re.exec("aaa123");
console.log(match[0]);
"#;
    assert_eq!(run_js(src), vec!["123"]);
}

#[test]
fn test_js_regexp_lookbehind_optional_prefix() {
    let src = r#"
const re = /(?<=(v)?)\d+/;
console.log(re.exec("v100")[0]);
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_regexp_nested_lookbehinds() {
    let src = r#"
const re = /(?<=(?<=\$)1)\d+/;
const match = re.exec("$150");
console.log(match[0]);
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_regexp_lookbehind_unicode_flag() {
    let src = r#"
const re = /(?<=\u{1F600})\w+/u;
const match = re.exec("😀happy");
console.log(match[0]);
"#;
    assert_eq!(run_js(src), vec!["happy"]);
}

#[test]
fn test_js_regexp_lookbehind_exec_null_on_failed_assertion() {
    let src = r#"
const re = /(?<=\$)(\d+)/;
console.log(re.exec("100 USD") === null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_negative_lookbehind_empty_match() {
    let src = r#"
const re = /(?<!a)/;
const match = re.exec("a");
console.log(match.index);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_regexp_lookbehind_with_indices_d_flag() {
    let src = r#"
const re = /(?<=\$)\d+/d;
const match = re.exec("Cost: $50");
console.log(match.indices[0].join(":")); // Indices match span of digits 50 (index 7..9)
"#;
    assert_eq!(run_js(src), vec!["7:9"]);
}

#[test]
fn test_js_regexp_lookbehind_multiline_flag() {
    let src = r#"
const re = /(?<=^\s*)\w+/m;
const match = re.exec("Line 1\n  Line 2");
console.log(match[0]);
"#;
    assert_eq!(run_js(src), vec!["Line"]);
}

#[test]
fn test_js_regexp_invalid_lookbehind_syntax_throws() {
    let src = r#"
try {
    eval("const re = /(?<=);/");
} catch (e) {
    console.log("Empty Lookbehind SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Empty Lookbehind SyntaxError"]);
}
