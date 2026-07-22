use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: RegExp Multiline (`m`), DotAll (`s`), IgnoreCase (`i`) & Global (`g`) Flags
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_regexp_dotall_s_flag_matches_newlines() {
    let src = r#"
const str = "a\nb";
console.log(`${/a.b/.test(str)}:${/a.b/s.test(str)}`);
"#;
    assert_eq!(run_js(src), vec!["false:true"]); // Dot '.' does NOT match newlines by default, but does with 's' flag!
}

#[test]
fn test_js_regexp_multiline_m_flag_caret_dollar_anchors() {
    let src = r#"
const str = "first\nsecond";
console.log(`${/^second/.test(str)}:${/^second/m.test(str)}`);
"#;
    assert_eq!(run_js(src), vec!["false:true"]);
}

#[test]
fn test_js_regexp_ignorecase_i_flag_case_insensitive() {
    let src = r#"
const str = "JavaScript";
console.log(`${/javascript/.test(str)}:${/javascript/i.test(str)}`);
"#;
    assert_eq!(run_js(src), vec!["false:true"]);
}

#[test]
fn test_js_regexp_global_g_flag_stateful_exec() {
    let src = r#"
const re = /a/g;
const str = "aba";
const m1 = re.exec(str);
const m2 = re.exec(str);
const m3 = re.exec(str);
console.log(`${m1.index}:${m2.index}:${m3 === null}`);
"#;
    assert_eq!(run_js(src), vec!["0:2:true"]);
}

#[test]
fn test_js_regexp_flags_property_order() {
    let src = r#"
const re = new RegExp("a", "migus");
console.log(re.flags); // Standard flags string ordering: "gimsuvy"!
"#;
    assert_eq!(run_js(src), vec!["gimsu"]);
}

#[test]
fn test_js_regexp_dotall_property() {
    let src = r#"
const re = /a/s;
console.log(re.dotAll + "|" + re.flags);
"#;
    assert_eq!(run_js(src), vec!["true|s"]);
}

#[test]
fn test_js_regexp_multiline_property() {
    let src = r#"
const re = /a/m;
console.log(re.multiline + "|" + re.flags);
"#;
    assert_eq!(run_js(src), vec!["true|m"]);
}

#[test]
fn test_js_regexp_ignorecase_property() {
    let src = r#"
const re = /a/i;
console.log(re.ignoreCase + "|" + re.flags);
"#;
    assert_eq!(run_js(src), vec!["true|i"]);
}

#[test]
fn test_js_regexp_global_property() {
    let src = r#"
const re = /a/g;
console.log(re.global + "|" + re.flags);
"#;
    assert_eq!(run_js(src), vec!["true|g"]);
}

#[test]
fn test_js_regexp_dotall_matches_carriage_return_and_line_feed() {
    let src = r#"
const str = "x\r\ny";
console.log(str.match(/x.*y/s)[0].length);
"#;
    assert_eq!(run_js(src), vec!["4"]);
}

#[test]
fn test_js_regexp_ignorecase_unicode_caseless_matching() {
    let src = r#"
const str = "Å"; // U+00C5
console.log(`${/å/i.test(str)}:${/å/iu.test(str)}`);
"#;
    assert_eq!(run_js(src), vec!["true:true"]);
}

#[test]
fn test_js_regexp_ignorecase_kelvin_symbol_caseless_matching() {
    let src = r#"
const str = "K";
console.log(/k/iu.test(str) + "|" + /k/iu.test("\u212A")); // Kelvin symbol \u212A matches 'k' with 'iu' flags!
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_regexp_duplicate_flags_throws_syntaxerror() {
    let src = r#"
try {
    eval("const re = /a/gg;");
} catch (e) {
    console.log("Duplicate Flags SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Duplicate Flags SyntaxError"]);
}

#[test]
fn test_js_regexp_invalid_flag_character_throws_syntaxerror() {
    let src = r#"
try {
    eval("const re = /a/z;");
} catch (e) {
    console.log("Invalid Flag SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Flag SyntaxError"]);
}

#[test]
fn test_js_regexp_constructor_flag_override() {
    let src = r#"
const re1 = /a/g;
const re2 = new RegExp(re1, "i"); // Passing flags to RegExp constructor overrides original pattern flags!
console.log(re2.flags);
"#;
    assert_eq!(run_js(src), vec!["i"]);
}

#[test]
fn test_js_regexp_multiline_m_flag_with_crlf() {
    let src = r#"
const str = "line1\r\nline2";
console.log(str.match(/^line2/m) !== null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_dotall_s_flag_with_line_separator_paragraph_separator() {
    let src = r#"
const str = "a\u2028b\u2029c"; // Line Separator U+2028 & Paragraph Separator U+2029
console.log(str.match(/a.*c/s) !== null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_regexp_global_g_flag_replace_all_matches() {
    let src = r#"
const str = "foo foo foo";
console.log(str.replace(/foo/g, "bar") + "|" + str.replace(/foo/, "bar"));
"#;
    assert_eq!(run_js(src), vec!["bar bar bar|bar foo foo"]);
}

#[test]
fn test_js_regexp_has_indices_d_flag_property() {
    let src = r#"
const re = /a/d;
console.log(re.hasIndices + "|" + re.flags);
"#;
    assert_eq!(run_js(src), vec!["true|d"]);
}

#[test]
fn test_js_regexp_flag_getters_on_prototype() {
    let src = r#"
const proto = RegExp.prototype;
console.log(`${proto.global}:${proto.ignoreCase}:${proto.multiline}:${proto.dotAll}`);
"#;
    assert_eq!(run_js(src), vec!["undefined:undefined:undefined:undefined"]);
}
