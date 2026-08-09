use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Tagged Template Literals & `strings.raw` Array Mechanics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_tagged_template_basic_string_array_and_expressions() {
    let src = r#"
function tag(strings, ...values) {
    return strings[0] + values[0] + strings[1] + values[1] + strings[2];
}
const a = 10, b = 20;
console.log(tag`X: ${a}, Y: ${b}!`);
"#;
    assert_eq!(run_js(src), vec!["X: 10, Y: 20!"]);
}

#[test]
fn test_js_tagged_template_raw_property_escapes() {
    let src = r#"
function rawTag(strings) {
    return strings.raw[0];
}
console.log(rawTag`Line1\nLine2\tTabbed`);
"#;
    assert_eq!(run_js(src), vec![r#"Line1\nLine2\tTabbed"#]);
}

#[test]
fn test_js_string_raw_builtin_tag() {
    let src = r#"
const path = String.raw`C:\Users\Name\Documents\file.txt`;
console.log(path);
"#;
    assert_eq!(run_js(src), vec![r#"C:\Users\Name\Documents\file.txt"#]);
}

#[test]
fn test_js_tagged_template_cooked_vs_raw_differences() {
    let src = r#"
function tag(strings) {
    return strings[0].length + "|" + strings.raw[0].length;
}
console.log(tag`\n`);
"#;
    assert_eq!(run_js(src), vec!["1|2"]);
}

#[test]
fn test_js_tagged_template_invalid_escape_sequence_cooked_undefined() {
    let src = r#"
function tag(strings) {
    return (strings[0] === undefined) + "|" + strings.raw[0];
}
console.log(tag`\unicode\xError`);
"#;
    assert_eq!(run_js(src), vec![r#"true|\unicode\xError"#]);
}

#[test]
fn test_js_tagged_template_frozen_strings_array() {
    let src = r#"
function tag(strings) {
    return Object.isFrozen(strings) + "|" + Object.isFrozen(strings.raw);
}
console.log(tag`Hello ${1}`);
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_tagged_template_html_escaping_sanitizer() {
    let src = r#"
function html(strings, ...values) {
    return strings.reduce((acc, str, i) => {
        let val = values[i - 1];
        if (typeof val === "string") {
            val = val.replace(/</g, "&lt;").replace(/>/g, "&gt;");
        }
        return acc + val + str;
    });
}
const user = "<script>alert(1)</script>";
console.log(html`<div>${user}</div>`);
"#;
    assert_eq!(
        run_js(src),
        vec!["<div>&lt;script&gt;alert(1)&lt;/script&gt;</div>"]
    );
}

#[test]
fn test_js_tagged_template_empty_string_literals() {
    let src = r#"
function inspect(strings, ...values) {
    return strings.length + "|" + values.length;
}
console.log(inspect``);
"#;
    assert_eq!(run_js(src), vec!["1|0"]);
}

#[test]
fn test_js_tagged_template_leading_trailing_interpolations() {
    let src = r#"
function tag(strings, ...values) {
    return strings.map(s => `"${s}"`).join("+");
}
console.log(tag`${1} middle ${2}`);
"#;
    assert_eq!(run_js(src), vec!["\"\"+ \" middle \" +\"\""]);
}

#[test]
fn test_js_string_raw_interleaved_interpolations() {
    let src = r#"
const name = "World";
console.log(String.raw`Hello ${name}\n!`);
"#;
    assert_eq!(run_js(src), vec![r#"Hello World\n!"#]);
}

#[test]
fn test_js_tagged_template_sql_parameterized_query_builder() {
    let src = r#"
function sql(strings, ...values) {
    const query = strings.join("?");
    return query + "|Params=" + values.join(",");
}
const id = 42, status = "active";
console.log(sql`SELECT * FROM users WHERE id = ${id} AND status = ${status}`);
"#;
    assert_eq!(
        run_js(src),
        vec!["SELECT * FROM users WHERE id = ? AND status = ?|Params=42,active"]
    );
}

#[test]
fn test_js_tagged_template_higher_order_tag_factory() {
    let src = r#"
function createPrefixTag(prefix) {
    return (strings, ...values) => {
        return prefix + strings[0] + values[0];
    };
}
const customTag = createPrefixTag("[LOG] ");
console.log(customTag`Value = ${100}`);
"#;
    assert_eq!(run_js(src), vec!["[LOG] Value = 100"]);
}

#[test]
fn test_js_tagged_template_this_context_method_call() {
    let src = r#"
const formatter = {
    prefix: ">>",
    tag(strings, ...values) {
        return `${this.prefix} ${strings[0]}${values[0]}`;
    }
};
console.log(formatter.tag`Val: ${50}`);
"#;
    assert_eq!(run_js(src), vec![">> Val: 50"]);
}

#[test]
fn test_js_tagged_template_symbol_values_handling() {
    let src = r#"
function tag(strings, val) {
    return strings[0] + String(val);
}
const sym = Symbol("symKey");
console.log(tag`Sym: ${sym}`);
"#;
    assert_eq!(run_js(src), vec!["Sym: Symbol(symKey)"]);
}

#[test]
fn test_js_tagged_template_bigint_values_handling() {
    let src = r#"
function tag(strings, val) {
    return strings[0] + val.toString();
}
console.log(tag`Big: ${1000n}`);
"#;
    assert_eq!(run_js(src), vec!["Big: 1000"]);
}

#[test]
fn test_js_tagged_template_multiline_raw_lines() {
    let src = r#"
function rawLines(strings) {
    return strings.raw[0].split("\n").length;
}
console.log(rawLines`Line 1
Line 2
Line 3`);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_tagged_template_iife_tag() {
    let src = r#"
const res = ((strings, ...values) => values[0] * 10)`Num: ${5}`;
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_tagged_template_member_expression_tag() {
    let src = r#"
const tags = {
    uppercase(strings, ...values) {
        return (strings[0] + values[0]).toUpperCase();
    }
};
console.log(tags.uppercase`hello ${"world"}`);
"#;
    assert_eq!(run_js(src), vec!["HELLO WORLD"]);
}

#[test]
fn test_js_tagged_template_subscript_expression_tag() {
    let src = r#"
const tagArray = [
    (strings, ...values) => values[0] + 1
];
console.log(tagArray[0]`${99}`);
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_string_raw_custom_object_emulation() {
    let src = r#"
const fakeStrings = { raw: ["A", "B"] };
console.log(String.raw(fakeStrings, 100));
"#;
    assert_eq!(run_js(src), vec!["A100B"]);
}

#[test]
fn test_js_string_raw_empty_template_literal() {
    let src = r#"
console.log(String.raw`` === "");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
