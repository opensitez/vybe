use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: RegExp Character Class Escapes (`\d`, `\D`, `\w`, `\W`, `\s`, `\S`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_regexp_digit_escape_d_and_D() {
    let src = r#"
const digits = "a1b2c3";
console.log(digits.match(/\d/g).join(",") + "|" + digits.match(/\D/g).join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3|a,b,c"]);
}

#[test]
fn test_js_regexp_word_character_escape_w_and_W() {
    let src = r#"
const str = "A_1 !@#";
console.log(str.match(/\w/g).join(",") + "|" + str.match(/\W/g).join(","));
"#;
    assert_eq!(run_js(src), vec!["A,_,1| ,!,@,#"]); // \w matches [a-zA-Z0-9_]
}

#[test]
fn test_js_regexp_whitespace_escape_s_and_S() {
    let src = r#"
const str = "A \t\nB";
console.log(str.match(/\s/g).length + "|" + str.match(/\S/g).join(","));
"#;
    assert_eq!(run_js(src), vec!["3|A,B"]); // \s matches spaces, tabs, newlines, form feeds
}

#[test]
fn test_js_regexp_character_class_custom_bracket() {
    let src = r#"
const str = "cat bat rat mat";
console.log(str.match(/[cbm]at/g).join(","));
"#;
    assert_eq!(run_js(src), vec!["cat,bat,mat"]);
}

#[test]
fn test_js_regexp_negated_character_class_bracket() {
    let src = r#"
const str = "cat bat rat mat";
console.log(str.match(/[^cbm]at/g).join(","));
"#;
    assert_eq!(run_js(src), vec!["rat"]);
}

#[test]
fn test_js_regexp_character_class_ranges() {
    let src = r#"
const str = "a5z9!";
console.log(str.match(/[a-z]/g).join(",") + "|" + str.match(/[0-9]/g).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,z|5,9"]);
}

#[test]
fn test_js_regexp_character_class_escape_inside_bracket() {
    let src = r#"
const str = "a 1 _ !";
console.log(str.match(/[\d\w]/g).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,1,_"]);
}

#[test]
fn test_js_regexp_backspace_escape_b_inside_character_class() {
    let src = r#"
const str = "a\b b";
console.log(str.match(/[\b]/g).length); // Inside character class [...], \b represents backspace (U+0008)!
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_regexp_digit_escape_with_unicode_flag() {
    let src = r#"
const str = "1₂3"; // ₂ is subscript digit U+2082
console.log(str.match(/\d/gu).join(",")); // ASCII digits 1 and 3 match \d!
"#;
    assert_eq!(run_js(src), vec!["1,3"]);
}

#[test]
fn test_js_regexp_word_escape_with_unicode_flag() {
    let src = r#"
const str = "a_1★";
console.log(str.match(/\w/gu).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,_,1"]);
}

#[test]
fn test_js_regexp_whitespace_escape_non_breaking_space() {
    let src = r#"
const str = "A\u00A0B"; // Non-breaking space \u00A0
console.log(str.match(/\s/g).length);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_regexp_whitespace_escape_byte_order_mark() {
    let src = r#"
const str = "A\uFEFFB"; // Zero-width non-breaking space (BOM) \uFEFF
console.log(str.match(/\s/g).length);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_regexp_hyphen_position_in_character_class() {
    let src = r#"
const str = "a-b+c";
console.log(str.match(/[-+]/g).join(",")); // Hyphen at start of character class is treated literally!
"#;
    assert_eq!(run_js(src), vec!["-,+"]);
}

#[test]
fn test_js_regexp_escaped_hyphen_in_character_class() {
    let src = r#"
const str = "a-b";
console.log(str.match(/[a\-z]/g).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,-"]);
}

#[test]
fn test_js_regexp_control_character_escape() {
    let src = r#"
const str = "line1\r\nline2";
console.log(str.match(/\r\n/g).length);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_regexp_hex_escape_x() {
    let src = r#"
const str = "ABC";
console.log(str.match(/\x41\x42/g).join(",")); // \x41 = 'A', \x42 = 'B'
"#;
    assert_eq!(run_js(src), vec!["AB"]);
}

#[test]
fn test_js_regexp_unicode_escape_u() {
    let src = r#"
const str = "Å";
console.log(str.match(/\u00C5/g).join(","));
"#;
    assert_eq!(run_js(src), vec!["Å"]);
}

#[test]
fn test_js_regexp_unicode_code_point_escape_u_braces() {
    let src = r#"
const str = "😀";
console.log(str.match(/\u{1F600}/gu).join(","));
"#;
    assert_eq!(run_js(src), vec!["😀"]);
}

#[test]
fn test_js_regexp_character_class_range_out_of_order_throws_syntaxerror() {
    let src = r#"
try {
    eval("const re = /[z-a]/;"); // Range out of order (z to a) is a SyntaxError!
} catch (e) {
    console.log("Range Out of Order SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Range Out of Order SyntaxError"]);
}

#[test]
fn test_js_regexp_digit_and_non_digit_complement() {
    let src = r##"
const str = "123abc";
console.log(str.replace(/\d/g, "#") + "|" + str.replace(/\D/g, "*"));
"##;
    assert_eq!(run_js(src), vec!["###abc|123***"]);
}

#[test]
fn test_js_regexp_constructor_digit_escape_string() {
    let src = r#"
const re = new RegExp("\\d+");
console.log(re.test("123"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
