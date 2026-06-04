/// Tagged template literals — tag functions, raw strings, cooked vs raw,
/// embedded expressions, nested tags, identity tag, HTML escaping patterns.
use super::helpers::run_js;

// ── basic tag function ────────────────────────────────────────────────────────

#[test]
fn tag_receives_strings_array_and_values() {
    assert_eq!(
        run_js(
            r#"
function tag(strings, ...values) {
    console.log(strings[0]);
    console.log(strings[1]);
    console.log(values[0]);
}
const x = 42;
tag`before${x}after`;
"#
        ),
        vec!["before", "after", "42"]
    );
}

#[test]
fn tag_can_reconstruct_template() {
    assert_eq!(
        run_js(
            r#"
function tag(strings, ...values) {
    return strings.reduce((acc, s, i) => acc + s + (values[i] !== undefined ? values[i] : ""), "");
}
const a = "hello";
const b = "world";
console.log(tag`${a}, ${b}!`);
"#
        ),
        vec!["hello, world!"]
    );
}

#[test]
fn tag_strings_length_is_values_plus_one() {
    assert_eq!(
        run_js(
            r#"
function tag(strings, ...values) {
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
fn tag_can_return_non_string() {
    assert_eq!(
        run_js(
            r#"
function tag(strings, ...values) {
    return values.reduce((a, b) => a + b, 0);
}
const result = tag`x${10}y${20}z${30}`;
console.log(result);
"#
        ),
        vec!["60"]
    );
}

// ── raw property ──────────────────────────────────────────────────────────────

#[test]
fn tag_raw_strings_preserve_backslash_n() {
    assert_eq!(
        run_js(
            r#"
function tag(strings) {
    return strings.raw[0].length;
}
const len = tag`line1\nline2`;
console.log(len);
"#
        ),
        vec!["12"]
    );
}

#[test]
fn tag_raw_different_from_cooked_for_escape() {
    assert_eq!(
        run_js(
            r#"
function tag(strings) {
    const cooked = strings[0];
    const raw = strings.raw[0];
    console.log(cooked !== raw);
}
tag`\n`;
"#
        ),
        vec!["true"]
    );
}

#[test]
fn string_raw_tag_preserves_backslash_sequences() {
    assert_eq!(
        run_js(
            r#"
const result = String.raw`Hello\nWorld`;
console.log(result.includes("\\n"));
console.log(result.includes("\n"));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn string_raw_with_interpolation() {
    assert_eq!(
        run_js(
            r#"
const path = String.raw`C:\Users\${" name ".trim()}`;
console.log(path);
"#
        ),
        vec!["C:\\Users\\name"]
    );
}

// ── tag transforms values ─────────────────────────────────────────────────────

#[test]
fn tag_can_uppercase_values() {
    assert_eq!(
        run_js(
            r#"
function upper(strings, ...values) {
    return strings.reduce((acc, s, i) => {
        const v = values[i] !== undefined ? String(values[i]).toUpperCase() : "";
        return acc + s + v;
    }, "");
}
const name = "world";
console.log(upper`hello ${name}!`);
"#
        ),
        vec!["hello WORLD!"]
    );
}

#[test]
fn tag_html_escape_pattern() {
    assert_eq!(
        run_js(
            r#"
function html(strings, ...values) {
    const escape = v => String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    return strings.reduce((acc, s, i) => acc + s + (values[i] !== undefined ? escape(values[i]) : ""), "");
}
const user = "<script>alert('xss')</script>";
console.log(html`<p>${user}</p>`);
"#
        ),
        vec!["<p>&lt;script&gt;alert('xss')&lt;/script&gt;</p>"]
    );
}

// ── tag with no interpolations ────────────────────────────────────────────────

#[test]
fn tag_no_expressions_gets_one_string() {
    assert_eq!(
        run_js(
            r#"
function tag(strings, ...values) {
    console.log(strings.length);
    console.log(values.length);
    console.log(strings[0]);
}
tag`just a string`;
"#
        ),
        vec!["1", "0", "just a string"]
    );
}

// ── nested tagged templates ───────────────────────────────────────────────────

#[test]
fn tag_within_tag_expression() {
    assert_eq!(
        run_js(
            r#"
function double(strings, ...values) {
    return values[0] * 2;
}
function outer(strings, ...values) {
    return "result:" + values[0];
}
console.log(outer`val=${double`${21}`}`);
"#
        ),
        vec!["result:42"]
    );
}

// ── identity tag ──────────────────────────────────────────────────────────────

#[test]
fn identity_tag_same_as_template_literal() {
    assert_eq!(
        run_js(
            r#"
function id(strings, ...values) {
    return String.raw({ raw: strings }, ...values);
}
const x = 5;
console.log(id`value is ${x}`);
"#
        ),
        vec!["value is 5"]
    );
}

// ── strings array is frozen ───────────────────────────────────────────────────

#[test]
fn tag_strings_array_is_reused_for_same_template() {
    assert_eq!(
        run_js(
            r#"
let firstRef;
function tag(strings) {
    if (!firstRef) firstRef = strings;
    return strings === firstRef;
}
function call() { return tag`hello`; }
const r1 = call();
const r2 = call();
console.log(r1);
console.log(r2);
"#
        ),
        vec!["true", "true"]
    );
}

// ── multi-line tagged templates ───────────────────────────────────────────────

#[test]
fn tag_preserves_newlines_in_string_parts() {
    assert_eq!(
        run_js(
            r#"
function tag(strings) {
    return strings[0].includes("\n");
}
console.log(tag`line1
line2`);
"#
        ),
        vec!["true"]
    );
}

// ── tag returning array ───────────────────────────────────────────────────────

#[test]
fn tag_can_return_array() {
    assert_eq!(
        run_js(
            r#"
function tokens(strings, ...values) {
    const result = [];
    strings.forEach((s, i) => {
        if (s) result.push(s);
        if (i < values.length) result.push(values[i]);
    });
    return result;
}
const a = 1, b = 2;
const parts = tokens`A${a}B${b}C`;
console.log(parts.join("-"));
"#
        ),
        vec!["A-1-B-2-C"]
    );
}
