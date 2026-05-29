/// String template tags, String.raw, raw vs cooked, tag returning non-string

use super::helpers::run_js;

#[test]
fn tagged_template_receives_strings_and_values() {
    assert_eq!(run_js(r#"
function tag(strings, ...values) {
    return strings.length + ":" + values.length;
}
const x = 1, y = 2;
console.log(tag`Hello ${x} world ${y}!`);
"#), vec!["3:2"]);
}

#[test]
fn tagged_template_cooked_array() {
    assert_eq!(run_js(r#"
function tag(strings) {
    return strings.join("-");
}
console.log(tag`a${1}b${2}c`);
"#), vec!["a-b-c"]);
}

#[test]
fn tagged_template_can_return_object() {
    assert_eq!(run_js(r#"
function tag(strings, ...vals) {
    return { strings, vals };
}
const result = tag`hello ${"world"}`;
console.log(result.strings[0]);
console.log(result.vals[0]);
"#), vec!["hello ", "world"]);
}

#[test]
fn string_raw_preserves_backslashes() {
    assert_eq!(run_js(r#"
const path = String.raw`C:\Users\test`;
console.log(path);
"#), vec!["C:\\Users\\test"]);
}

#[test]
fn string_raw_preserves_newline_sequence() {
    assert_eq!(run_js(r#"
const s = String.raw`line1\nline2`;
console.log(s.includes("\\n"));
console.log(s.includes("\n"));
"#), vec!["true", "false"]);
}

#[test]
fn raw_property_on_strings_array() {
    assert_eq!(run_js(r#"
function tag(strings) {
    return strings.raw[0];
}
const result = tag`\n\t`;
console.log(result); // raw: \\n\\t
console.log(result.length);
"#), vec!["\\n\\t", "4"]);
}

#[test]
fn tagged_template_html_escape() {
    assert_eq!(run_js(r#"
function html(strings, ...values) {
    return strings.reduce((acc, str, i) => {
        const val = values[i - 1] != null
            ? String(values[i - 1]).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
            : "";
        return acc + val + str;
    });
}
const user = "<script>alert(1)</script>";
const result = html`<p>Hello ${user}!</p>`;
console.log(result.includes("&lt;script&gt;"));
"#), vec!["true"]);
}

#[test]
fn tagged_template_sql_builder() {
    assert_eq!(run_js(r#"
function sql(strings, ...values) {
    const query = strings.reduce((acc, s, i) => acc + s + (i < values.length ? "?" : ""), "");
    return { query, params: values };
}
const id = 42, name = "Alice";
const result = sql`SELECT * FROM users WHERE id = ${id} AND name = ${name}`;
console.log(result.params.length);
console.log(result.params[0]);
console.log(result.params[1]);
"#), vec!["2", "42", "Alice"]);
}

#[test]
fn nested_tagged_template() {
    assert_eq!(run_js(r#"
const tag = (s, ...v) => s.join("") + v.join("");
const inner = tag`a${1}b`;
const result = tag`x${inner}y`;
console.log(result);
"#), vec!["xyab1"]);
}

#[test]
fn tagged_template_with_expression_values() {
    assert_eq!(run_js(r#"
const id = (s, ...v) => String.raw({ raw: s }, ...v);
const a = 3, b = 4;
const result = id`${a} + ${b} = ${a + b}`;
console.log(result);
"#), vec!["3 + 4 = 7"]);
}
