/// Template literals advanced — complex expressions, tagged templates with arrays,
/// nesting, template string formatting patterns, heredoc-style, error in template.
use super::helpers::run_js;

// ── expression interpolation ──────────────────────────────────────────────────

#[test]
fn template_ternary_expression() {
    assert_eq!(
        run_js(
            r#"
const x = 5;
console.log(`x is ${x > 3 ? "big" : "small"}`);
"#
        ),
        vec!["x is big"]
    );
}

#[test]
fn template_function_call_in_expression() {
    assert_eq!(
        run_js(
            r#"
function double(n) { return n * 2; }
console.log(`result: ${double(21)}`);
"#
        ),
        vec!["result: 42"]
    );
}

#[test]
fn template_object_method_call() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3];
console.log(`items: ${arr.join(", ")}`);
"#
        ),
        vec!["items: 1, 2, 3"]
    );
}

// ── nesting templates ─────────────────────────────────────────────────────────

#[test]
fn nested_template_literals() {
    assert_eq!(
        run_js(
            r#"
const items = ["a", "b", "c"];
const result = `list: ${items.map(i => `[${i}]`).join(",")}`;
console.log(result);
"#
        ),
        vec!["list: [a],[b],[c]"]
    );
}

#[test]
fn nested_template_in_condition() {
    assert_eq!(
        run_js(
            r#"
const n = 3;
const msg = `${n} item${n !== 1 ? `s` : ""}`;
console.log(msg);
"#
        ),
        vec!["3 items"]
    );
}

// ── multiline templates ───────────────────────────────────────────────────────

#[test]
fn multiline_template_preserves_newlines() {
    assert_eq!(
        run_js(
            r#"
const text = `line1
line2
line3`;
const lines = text.split("\n");
console.log(lines.length);
console.log(lines[1]);
"#
        ),
        vec!["3", "line2"]
    );
}

// ── tagged templates ──────────────────────────────────────────────────────────

#[test]
fn tagged_template_strings_array_length() {
    assert_eq!(
        run_js(
            r#"
function tag(strings, ...values) {
    // strings.length is always values.length + 1
    console.log(strings.length);
    console.log(values.length);
}
tag`a${1}b${2}c`;
"#
        ),
        vec!["3", "2"]
    );
}

#[test]
fn tagged_template_reconstruct_original() {
    assert_eq!(
        run_js(
            r#"
function identity(strings, ...values) {
    return strings.reduce((acc, str, i) => acc + (values[i-1] ?? "") + str);
}
const x = 42;
console.log(identity`value is ${x} done`);
"#
        ),
        vec!["value is 42 done"]
    );
}

#[test]
fn tagged_template_transform_values() {
    assert_eq!(
        run_js(
            r#"
function upper(strings, ...values) {
    return strings.reduce((acc, str, i) => {
        const val = values[i-1];
        return acc + (val !== undefined ? String(val).toUpperCase() : "") + str;
    });
}
const name = "alice";
console.log(upper`Hello, ${name}!`);
"#
        ),
        vec!["Hello, ALICE!"]
    );
}

#[test]
fn tagged_template_html_escape() {
    assert_eq!(
        run_js(
            r#"
function html(strings, ...values) {
    function escape(s) {
        return String(s)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");
    }
    return strings.reduce((acc, str, i) =>
        acc + (i > 0 ? escape(values[i-1]) : "") + str
    );
}
const user = "<script>alert(1)</script>";
console.log(html`Hello ${user}!`);
"#
        ),
        vec!["Hello &lt;script&gt;alert(1)&lt;/script&gt;!"]
    );
}

// ── String.raw ────────────────────────────────────────────────────────────────

#[test]
fn string_raw_preserves_escape_sequences() {
    assert_eq!(
        run_js(
            r#"
const path = String.raw`C:\Users\name\file.txt`;
console.log(path);
"#
        ),
        vec!["C:\\Users\\name\\file.txt"]
    );
}

#[test]
fn string_raw_vs_cooked() {
    assert_eq!(
        run_js(
            r#"
const raw = String.raw`\n\t`;
const cooked = `\n\t`;
console.log(raw.length);   // 4 chars: \, n, \, t
console.log(cooked.length); // 2 chars: newline, tab
"#
        ),
        vec!["4", "2"]
    );
}

// ── template as heredoc ───────────────────────────────────────────────────────

#[test]
fn template_as_heredoc_style() {
    assert_eq!(
        run_js(
            r#"
function dedent(str) {
    const lines = str.split("\n").filter(l => l.trim());
    const indent = Math.min(...lines.map(l => l.match(/^\s*/)[0].length));
    return lines.map(l => l.slice(indent)).join("\n");
}
const code = dedent(`
    function hello() {
        return "world";
    }
`);
console.log(code.startsWith("function"));
"#
        ),
        vec!["true"]
    );
}

// ── template with complex object interpolation ────────────────────────────────

#[test]
fn template_with_object_tostring() {
    assert_eq!(
        run_js(
            r#"
const obj = { toString() { return "custom"; } };
console.log(`value: ${obj}`);
"#
        ),
        vec!["value: custom"]
    );
}

#[test]
fn template_with_null_and_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(`${null}`);
console.log(`${undefined}`);
console.log(`${false}`);
console.log(`${0}`);
"#
        ),
        vec!["null", "undefined", "false", "0"]
    );
}

// ── recursive/function template ───────────────────────────────────────────────

#[test]
fn template_in_reduce_pattern() {
    assert_eq!(
        run_js(
            r#"
const items = [{ name: "a", val: 1 }, { name: "b", val: 2 }];
const html = items.map(({ name, val }) => `${name}=${val}`).join(", ");
console.log(html);
"#
        ),
        vec!["a=1, b=2"]
    );
}

// ── tagged template with raw property ─────────────────────────────────────────

#[test]
fn tagged_template_raw_strings_available() {
    assert_eq!(
        run_js(
            r#"
function raw(strings) {
    return strings.raw[0];
}
const result = raw`\n\t`;
console.log(result.length); // raw: 4 chars
console.log(result[0]);     // backslash
"#
        ),
        vec!["4", "\\"]
    );
}
