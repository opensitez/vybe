use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: RegExp `v` Flag Unicode Sets & Character Class Set Operations (ES2024)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_regexp_v_flag_unicode_sets_property() {
    let src = r#"
const re = /[\p{Decimal_Number}]/v;
console.log(re.unicodeSets + "|" + re.flags);
"#;
    assert_eq!(run_js(src), vec!["true|v"]);
}

#[test]
fn test_js_regexp_v_flag_string_disjunction_q_sequence() {
    let src = r#"
const re = /[\q{abc|def}]/v;
console.log(re.test("abc") + "|" + re.test("def") + "|" + re.test("ab"));
"#;
    assert_eq!(run_js(src), vec!["true|true|false"]);
}

#[test]
fn test_js_regexp_v_flag_set_intersection_ampersands() {
    let src = r#"
const re = /[\p{ASCII}&&[a-z]]/v;
console.log(re.test("a") + "|" + re.test("A") + "|" + re.test("5"));
"#;
    assert_eq!(run_js(src), vec!["true|false|false"]);
}

#[test]
fn test_js_regexp_v_flag_set_subtraction_dashes() {
    let src = r#"
const re = /[[a-z]--[aeiou]]/v; // Consonants only!
console.log(re.test("b") + "|" + re.test("a") + "|" + re.test("e"));
"#;
    assert_eq!(run_js(src), vec!["true|false|false"]);
}

#[test]
fn test_js_regexp_v_flag_nested_class_brackets() {
    let src = r#"
const re = /[a-z[0-9]]/v;
console.log(re.test("x") + "|" + re.test("5") + "|" + re.test("$"));
"#;
    assert_eq!(run_js(src), vec!["true|true|false"]);
}

#[test]
fn test_js_regexp_v_flag_emoji_sequence_matching() {
    let src = r#"
const re = /^\p{RGI_Emoji}$/v;
console.log(re.test("😀") + "|" + re.test("⚽") + "|" + re.test("A"));
"#;
    assert_eq!(run_js(src), vec!["true|true|false"]);
}

#[test]
fn test_js_regexp_v_flag_mutually_exclusive_with_u_flag_throws_syntaxerror() {
    let src = r#"
try {
    eval("const re = /test/uv;"); // Flags u and v cannot be combined!
} catch (e) {
    console.log("u and v flags SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["u and v flags SyntaxError"]);
}

#[test]
fn test_js_regexp_v_flag_set_intersection_multiple() {
    let src = r#"
const re = /[a-z&&[b-z]&&[c-z]]/v;
console.log(re.test("c") + "|" + re.test("b") + "|" + re.test("a"));
"#;
    assert_eq!(run_js(src), vec!["true|false|false"]);
}

#[test]
fn test_js_regexp_v_flag_set_subtraction_multiple() {
    let src = r#"
const re = /[a-z--[a]--[b]]/v;
console.log(re.test("c") + "|" + re.test("a") + "|" + re.test("b"));
"#;
    assert_eq!(run_js(src), vec!["true|false|false"]);
}

#[test]
fn test_js_regexp_v_flag_empty_q_sequence() {
    let src = r#"
const re = /[\q{}]/v;
console.log(re.test(""));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_v_flag_negated_class_with_string_disjunction_throws_syntaxerror() {
    let src = r#"
try {
    eval("const re = /[^\\q{abc}]/v;"); // Negated character classes cannot contain string sequences!
} catch (e) {
    console.log("Negated Sequence Class SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Negated Sequence Class SyntaxError"]);
}

#[test]
fn test_js_regexp_v_flag_unicode_script_extensions() {
    let src = r#"
const re = /\p{Script_Extensions=Greek}/v;
console.log(re.test("α") + "|" + re.test("a"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_regexp_v_flag_escaped_special_chars_in_sets() {
    let src = r#"
const re = /[\(\)\[\]\{\}]/v;
console.log(re.test("(") + "|" + re.test("[") + "|" + re.test("a"));
"#;
    assert_eq!(run_js(src), vec!["true|true|false"]);
}

#[test]
fn test_js_regexp_v_flag_multiline_and_global_combo() {
    let src = r#"
const re = /[\p{ASCII_Hex_Digit}]/gv;
const matches = [..."a1z".matchAll(re)];
console.log(matches.map(m => m[0]).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,1"]);
}

#[test]
fn test_js_regexp_v_flag_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(RegExp.prototype, "unicodeSets");
console.log(typeof desc.get === "function");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_v_flag_ignore_case() {
    let src = r#"
const re = /[[a-z]--[a]]/iv;
console.log(re.test("B") + "|" + re.test("A"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_regexp_v_flag_exec_match_indices() {
    let src = r#"
const re = /[\q{hello}]/dv;
const match = re.exec("hello world");
console.log(match.indices[0].join(","));
"#;
    assert_eq!(run_js(src), vec!["0,5"]);
}

#[test]
fn test_js_regexp_v_flag_surrogate_pair_handling() {
    let src = r#"
const re = /[😀-🎉]/v;
console.log(re.test("😀") + "|" + re.test("🎈") + "|" + re.test("A"));
"#;
    assert_eq!(run_js(src), vec!["true|true|false"]);
}

#[test]
fn test_js_regexp_v_flag_set_subtraction_with_intersection() {
    let src = r#"
const re = /[[a-z]--[a-m&&[c-g]]]/v; // Excludes c, d, e, f, g
console.log(re.test("a") + "|" + re.test("b") + "|" + re.test("c") + "|" + re.test("g"));
"#;
    assert_eq!(run_js(src), vec!["true|true|false|false"]);
}

#[test]
fn test_js_regexp_v_flag_constructor_string_pattern() {
    let src = r#"
const re = new RegExp("[\\p{ASCII}&&[0-9]]", "v");
console.log(re.test("5") + "|" + re.test("x"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}
