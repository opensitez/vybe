use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Multiline Template Literals & Escape Sequences
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_multiline_template_literal_newline_preservation() {
    let src = r#"
const multiline = `Line1
Line2
Line3`;
console.log(multiline.split("\n").length);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_multiline_template_literal_crlf_normalization() {
    let src = r#"
const str = `A
B`;
console.log(str.length);
"#;
    assert_eq!(run_js(src), vec!["3"]); // "A\nB" length is 3
}

#[test]
fn test_js_template_literal_tab_character_preservation() {
    let src = r#"
const tabbed = `Col1\tCol2`;
console.log(tabbed.includes("\t"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_template_literal_hex_escape_sequences() {
    let src = r#"
const hex = `\x48\x65\x6c\x6c\x6f`;
console.log(hex);
"#;
    assert_eq!(run_js(src), vec!["Hello"]);
}

#[test]
fn test_js_template_literal_unicode_escape_sequences() {
    let src = r#"
const unicode = `\u0041\u0042\u0043`;
console.log(unicode);
"#;
    assert_eq!(run_js(src), vec!["ABC"]);
}

#[test]
fn test_js_template_literal_unicode_codepoint_escapes() {
    let src = r#"
const emoji = `\u{1F600}`;
console.log(emoji.codePointAt(0).toString(16));
"#;
    assert_eq!(run_js(src), vec!["1f600"]);
}

#[test]
fn test_js_template_literal_escaped_backtick() {
    let src = r#"
const str = `Backtick: \``;
console.log(str);
"#;
    assert_eq!(run_js(src), vec!["Backtick: `"]);
}

#[test]
fn test_js_template_literal_escaped_dollar_curly() {
    let src = r#"
const val = 100;
console.log(`Literal: \${val}`);
"#;
    assert_eq!(run_js(src), vec!["Literal: ${val}"]);
}

#[test]
fn test_js_template_literal_escaped_backslash() {
    let src = r#"
const path = `C:\\Program Files\\App`;
console.log(path);
"#;
    assert_eq!(run_js(src), vec![r#"C:\Program Files\App"#]);
}

#[test]
fn test_js_multiline_template_literal_indentation_preservation() {
    let src = r#"
const indented = `  Indent1
    Indent2`;
console.log(indented.startsWith("  "));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_template_literal_null_character_escape() {
    let src = r#"
const nullChar = `\0`;
console.log(nullChar.charCodeAt(0));
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_template_literal_vertical_tab_and_form_feed() {
    let src = r#"
const str = `\v\f`;
console.log(str.charCodeAt(0) + "|" + str.charCodeAt(1));
"#;
    assert_eq!(run_js(src), vec!["11|12"]);
}

#[test]
fn test_js_template_literal_octal_escapes_prohibited_in_tagged_cooked() {
    let src = r#"
function tag(strings) {
    return strings[0] === undefined; // Cooked is undefined for non-octal legacy escapes in ES2018
}
console.log(tag`\0123`);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_multiline_template_literal_html_template_string() {
    let src = r#"
const title = "Dashboard";
const html = `
<div class="container">
    <h1>${title}</h1>
</div>`.trim();
console.log(html.startsWith("<div"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_template_literal_line_continuation_escape() {
    let src = r#"
const singleLine = `Line 1 \
Line 2`;
console.log(singleLine.includes("\n"));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_template_literal_carriage_return_escape() {
    let src = r#"
const str = `CR:\rEnd`;
console.log(str.includes("\r"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_multiline_template_literal_empty_lines() {
    let src = r#"
const str = `A

B`;
console.log(str.split("\n").length);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_template_literal_unicode_surrogate_pair_escapes() {
    let src = r#"
const surrogatePair = `\uD83D\uDE00`;
console.log(surrogatePair.codePointAt(0).toString(16));
"#;
    assert_eq!(run_js(src), vec!["1f600"]);
}

#[test]
fn test_js_template_literal_escaped_single_and_double_quotes() {
    let src = r#"
const str = `Quotes: \' and \"`;
console.log(str);
"#;
    assert_eq!(run_js(src), vec![r#"Quotes: ' and ""#]);
}

#[test]
fn test_js_template_literal_trailing_newline() {
    let src = r#"
const str = `Header
`;
console.log(str.endsWith("\n"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_template_literal_invalid_hex_escape_in_tagged_cooked_undefined() {
    let src = r#"
function tag(strings) {
    return strings[0] === undefined;
}
console.log(tag`\xZZ`);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

