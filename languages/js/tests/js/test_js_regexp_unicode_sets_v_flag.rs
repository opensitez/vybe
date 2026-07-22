use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: RegExp Unicode Sets (`/v` flag - ES2024 Set Operations in Character Classes)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_regexp_unicode_sets_v_flag_property() {
    let src = r#"
const re = /[\p{Script=Greek}]/v;
console.log(re.unicodeSets + "|" + re.unicode);
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_regexp_unicode_sets_u_and_v_mutually_exclusive() {
    let src = r#"
try {
    eval("const re = /a/uv;");
} catch (e) {
    console.log("u and v Flags Mutually Exclusive SyntaxError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["u and v Flags Mutually Exclusive SyntaxError"]
    );
}

#[test]
fn test_js_regexp_unicode_sets_difference_operator() {
    let src = r#"
const re = /[\p{Decimal_Number}--[0-4]]/v; // Digits except 0..4
console.log(re.test("5") + "|" + re.test("3"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_regexp_unicode_sets_intersection_operator() {
    let src = r#"
const re = /[\p{ASCII}&&[\p{Letter}]]/v; // ASCII letters
console.log(re.test("A") + "|" + re.test("5") + "|" + re.test("α"));
"#;
    assert_eq!(run_js(src), vec!["true|false|false"]);
}

#[test]
fn test_js_regexp_unicode_sets_string_disjunction_properties() {
    let src = r#"
const re = /[\p{Basic_Emoji}]/v; // Matches multi-character emoji strings
console.log(re.test("😀"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_unicode_sets_nested_character_classes() {
    let src = r#"
const re = /[[a-z]--[aeiou]]/v; // Consonants only
console.log(re.test("b") + "|" + re.test("e"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_regexp_unicode_sets_string_literals_in_class() {
    let src = r#"
const re = /[a|b|{hello}]/v; // Matches "a", "b", or string "hello"
console.log(re.test("a") + "|" + re.test("hello"));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_regexp_unicode_sets_negated_class_with_strings_throws() {
    let src = r#"
try {
    eval("const re = /[^\\p{Basic_Emoji}]/v;");
} catch (e) {
    console.log("Negated String Property Class SyntaxError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Negated String Property Class SyntaxError"]
    );
}

#[test]
fn test_js_regexp_unicode_sets_flags_accessor_includes_v() {
    let src = r#"
const re = new RegExp("abc", "vgi");
console.log(re.flags);
"#;
    assert_eq!(run_js(src), vec!["giv"]);
}

#[test]
fn test_js_regexp_unicode_sets_greek_script_matching() {
    let src = r#"
const re = /^\p{Script=Greek}+$/v;
console.log(re.test("αβγ") + "|" + re.test("abc"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_regexp_unicode_sets_emoji_sequence_matching() {
    let src = r#"
const re = /^\p{RGI_Emoji}+$/v;
console.log(re.test("👍") + "|" + re.test("A"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_regexp_unicode_sets_intersection_chaining() {
    let src = r#"
const re = /[a-z]&&[c-z]&&[a-f]/v; // Elements in a-z AND c-z AND a-f -> c,d,e,f
console.log(re.test("d") + "|" + re.test("a") + "|" + re.test("z"));
"#;
    assert_eq!(run_js(src), vec!["true|false|false"]);
}

#[test]
fn test_js_regexp_unicode_sets_subtraction_chaining() {
    let src = r#"
const re = /[a-z]--[a-m]--[x-z]/v; // Elements n..w
console.log(re.test("p") + "|" + re.test("b") + "|" + re.test("z"));
"#;
    assert_eq!(run_js(src), vec!["true|false|false"]);
}

#[test]
fn test_js_regexp_unicode_sets_escaped_operators_in_class() {
    let src = r#"
const re = /[\--\-]/v; // Range from '-' to '-'
console.log(re.test("-"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_unicode_sets_case_insensitive_flag_v() {
    let src = r#"
const re = /[\p{Lower}]/vi;
console.log(re.test("A"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_unicode_sets_code_point_match() {
    let src = r#"
const re = /^\u{1F600}$/v;
console.log(re.test("😀"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_unicode_sets_curly_brace_literal_requires_escape_without_v() {
    let src = r#"
const reV = /[a{b}]/v; // Valid in v-mode (matches "a" or string "{b}")
console.log(reV.test("{b}"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_unicode_sets_class_subtraction_with_emoji() {
    let src = r#"
const re = /[\p{Emoji}--[0-9#*]]/v;
console.log(re.test("😀") + "|" + re.test("1"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_regexp_unicode_sets_constructor_flags_parsing() {
    let src = r#"
const re = new RegExp("\\p{Letter}", "v");
console.log(re.unicodeSets);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_unicode_sets_syntax_error_mixing_operators_without_grouping() {
    let src = r#"
try {
    eval("const re = /[a-z]&&[0-9]--[a]/v;");
} catch (e) {
    console.log("Mixed Operators Without Grouping SyntaxError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Mixed Operators Without Grouping SyntaxError"]
    );
}
