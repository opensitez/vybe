use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: RegExp Quantifiers (`*`, `+`, `?`, `{n,m}`) & Greedy vs Lazy Matching
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_regexp_greedy_star_quantifier() {
    let src = r#"
const html = "<div>first</div><div>second</div>";
console.log(html.match(/<div>.*<\/div>/)[0]);
"#;
    assert_eq!(run_js(src), vec!["<div>first</div><div>second</div>"]);
}

#[test]
fn test_js_regexp_lazy_star_quantifier() {
    let src = r#"
const html = "<div>first</div><div>second</div>";
console.log(html.match(/<div>.*?<\/div>/g).join("|"));
"#;
    assert_eq!(run_js(src), vec!["<div>first</div>|<div>second</div>"]);
}

#[test]
fn test_js_regexp_greedy_plus_quantifier() {
    let src = r#"
const str = "aaaaa";
console.log(str.match(/a+/)[0]);
"#;
    assert_eq!(run_js(src), vec!["aaaaa"]);
}

#[test]
fn test_js_regexp_lazy_plus_quantifier() {
    let src = r#"
const str = "aaaaa";
console.log(str.match(/a+?/)[0]);
"#;
    assert_eq!(run_js(src), vec!["a"]);
}

#[test]
fn test_js_regexp_greedy_optional_quantifier() {
    let src = r#"
const str = "color colour";
console.log(str.match(/colou?r/g).join(","));
"#;
    assert_eq!(run_js(src), vec!["color,colour"]);
}

#[test]
fn test_js_regexp_lazy_optional_quantifier() {
    let src = r#"
const str = "aaaaa";
console.log(str.match(/a??/)[0]); // Matches 0 occurrences!
"#;
    assert_eq!(run_js(src), vec![""]);
}

#[test]
fn test_js_regexp_greedy_range_quantifier() {
    let src = r#"
const str = "123456789";
console.log(str.match(/\d{2,5}/)[0]);
"#;
    assert_eq!(run_js(src), vec!["12345"]);
}

#[test]
fn test_js_regexp_lazy_range_quantifier() {
    let src = r#"
const str = "123456789";
console.log(str.match(/\d{2,5}?/)[0]);
"#;
    assert_eq!(run_js(src), vec!["12"]);
}

#[test]
fn test_js_regexp_exact_range_quantifier() {
    let src = r#"
const str = "123456";
console.log(str.match(/\d{4}/)[0]);
"#;
    assert_eq!(run_js(src), vec!["1234"]);
}

#[test]
fn test_js_regexp_open_ended_range_quantifier() {
    let src = r#"
const str = "1234567";
console.log(str.match(/\d{3,}/)[0]);
"#;
    assert_eq!(run_js(src), vec!["1234567"]);
}

#[test]
fn test_js_regexp_lazy_open_ended_range_quantifier() {
    let src = r#"
const str = "1234567";
console.log(str.match(/\d{3,}?/)[0]);
"#;
    assert_eq!(run_js(src), vec!["123"]);
}

#[test]
fn test_js_regexp_quantifier_on_grouped_expression() {
    let src = r#"
const str = "ha-ha-ha";
console.log(str.match(/(ha-)+/)[0]);
"#;
    assert_eq!(run_js(src), vec!["ha-ha-"]);
}

#[test]
fn test_js_regexp_quantifier_on_non_capturing_group() {
    let src = r#"
const str = "ha-ha-ha";
console.log(str.match(/(?:ha-)+/)[0]);
"#;
    assert_eq!(run_js(src), vec!["ha-ha-"]);
}

#[test]
fn test_js_regexp_invalid_range_min_greater_than_max_throws_syntaxerror() {
    let src = r#"
try {
    eval("const re = /a{5,2}/;"); // min 5 > max 2 is a SyntaxError!
} catch (e) {
    console.log("Invalid Quantifier Range SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Quantifier Range SyntaxError"]);
}

#[test]
fn test_js_regexp_nested_quantifiers() {
    let src = r#"
const str = "aaa bbb";
console.log(str.match(/(?:a+)+/)[0]);
"#;
    assert_eq!(run_js(src), vec!["aaa"]);
}

#[test]
fn test_js_regexp_zero_quantifier_match_at_every_position() {
    let src = r#"
const str = "ab";
console.log(str.match(/a*/g).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,,"]); // Matches 'a' at 0, '' at 1, '' at 2
}

#[test]
fn test_js_regexp_quantifier_surrogate_pair_with_u_flag() {
    let src = r#"
const str = "😀😀😀";
console.log(str.match(/😀{2}/u)[0]);
"#;
    assert_eq!(run_js(src), vec!["😀😀"]);
}

#[test]
fn test_js_regexp_quantifier_without_target_throws_syntaxerror() {
    let src = r#"
try {
    eval("const re = /*/;");
} catch (e) {
    console.log("Dangling Quantifier SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Dangling Quantifier SyntaxError"]);
}

#[test]
fn test_js_regexp_lazy_quantifier_backtracking() {
    let src = r#"
const str = "axxxb";
console.log(str.match(/a.*?b/)[0]);
"#;
    assert_eq!(run_js(src), vec!["axxxb"]); // Lazy .*? expands as needed to satisfy trailing 'b'!
}

#[test]
fn test_js_regexp_greedy_quantifier_backtracking() {
    let src = r#"
const str = "axxxb";
console.log(str.match(/a.*b/)[0]);
"#;
    assert_eq!(run_js(src), vec!["axxxb"]);
}
