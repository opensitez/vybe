use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Strings — template literals, methods, patterns
// ═══════════════════════════════════════════════════════════

// ── Template literals ──────────────────────────────────────

#[test]
fn template_literal_basic() {
    let out = run_js(
        r#"
const name = "World";
console.log(`Hello ${name}!`);
"#,
    );
    assert_eq!(out, vec!["Hello World!"]);
}

#[test]
fn template_literal_expression() {
    let out = run_js(
        r#"
const a = 5, b = 10;
console.log(`Sum: ${a + b}`);
"#,
    );
    assert_eq!(out, vec!["Sum: 15"]);
}

#[test]
fn template_literal_multiline() {
    let out = run_js(
        r#"
const text = `line1
line2`;
console.log(text);
"#,
    );
    assert_eq!(out, vec!["line1\nline2"]);
}

#[test]
fn template_literal_nested() {
    let out = run_js(
        r#"
const x = 5;
console.log(`result: ${x > 3 ? `big(${x})` : "small"}`);
"#,
    );
    assert_eq!(out, vec!["result: big(5)"]);
}

// ── String methods ─────────────────────────────────────────

#[test]
fn string_length() {
    let out = run_js(r#"console.log("hello".length);"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn string_charat() {
    let out = run_js(r#"console.log("hello".charAt(1));"#);
    assert_eq!(out, vec!["e"]);
}

#[test]
fn string_indexof() {
    let out = run_js(
        r#"
console.log("hello world".indexOf("world"));
console.log("hello world".indexOf("xyz"));
"#,
    );
    assert_eq!(out, vec!["6", "-1"]);
}

#[test]
fn string_includes() {
    let out = run_js(
        r#"
console.log("hello world".includes("world"));
console.log("hello world".includes("xyz"));
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn string_startswith() {
    let out = run_js(
        r#"
console.log("hello".startsWith("hel"));
console.log("hello".startsWith("ell"));
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn string_endswith() {
    let out = run_js(
        r#"
console.log("hello".endsWith("llo"));
console.log("hello".endsWith("hel"));
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn string_slice() {
    let out = run_js(
        r#"
console.log("hello world".slice(6));
console.log("hello world".slice(0, 5));
"#,
    );
    assert_eq!(out, vec!["world", "hello"]);
}

#[test]
fn string_substring() {
    let out = run_js(
        r#"
console.log("hello".substring(1, 4));
"#,
    );
    assert_eq!(out, vec!["ell"]);
}

#[test]
fn string_touppercase() {
    let out = run_js(r#"console.log("hello".toUpperCase());"#);
    assert_eq!(out, vec!["HELLO"]);
}

#[test]
fn string_tolowercase() {
    let out = run_js(r#"console.log("HELLO".toLowerCase());"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn string_trim() {
    let out = run_js(r#"console.log("  hello  ".trim());"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn string_split() {
    let out = run_js(
        r#"
const parts = "a,b,c".split(",");
console.log(parts.length);
console.log(parts.join("|"));
"#,
    );
    assert_eq!(out, vec!["3", "a|b|c"]);
}

#[test]
fn string_replace() {
    let out = run_js(
        r#"
console.log("hello world".replace("world", "JS"));
"#,
    );
    assert_eq!(out, vec!["hello JS"]);
}

#[test]
fn string_repeat() {
    let out = run_js(r#"console.log("ab".repeat(3));"#);
    assert_eq!(out, vec!["ababab"]);
}

#[test]
fn string_padstart() {
    let out = run_js(r#"console.log("5".padStart(3, "0"));"#);
    assert_eq!(out, vec!["005"]);
}

#[test]
fn string_padend() {
    let out = run_js(r#"console.log("5".padEnd(3, "0"));"#);
    assert_eq!(out, vec!["500"]);
}

#[test]
fn string_concat() {
    let out = run_js(
        r#"
console.log("hello" + " " + "world");
console.log("num: " + 42);
console.log(1 + "2");
"#,
    );
    assert_eq!(out, vec!["hello world", "num: 42", "12"]);
}

#[test]
fn string_comparison() {
    let out = run_js(
        r#"
console.log("a" < "b");
console.log("b" > "a");
console.log("abc" === "abc");
"#,
    );
    assert_eq!(out, vec!["true", "true", "true"]);
}

#[test]
fn string_slice_negative_index() {
    let out = run_js(
        r#"
console.log("abcdef".slice(-2));
console.log("abcdef".slice(-3, -1));
"#,
    );
    assert_eq!(out, vec!["ef", "de"]);
}

#[test]
fn string_trim_start_end() {
    let out = run_js(
        r#"
const s = "\t hello \n";
console.log(s.trimStart() === "hello \n");
console.log(s.trimEnd() === "\t hello");
console.log(s.trim());
"#,
    );
    assert_eq!(out, vec!["true", "true", "hello"]);
}

#[test]
fn string_at_and_index_access() {
    let out = run_js(
        r#"
console.log("abc".at(1));
console.log("abc".at(-1));
console.log("".at(0));
console.log("abc"[2]);
"#,
    );
    assert_eq!(out, vec!["b", "c", "undefined", "c"]);
}

#[test]
fn string_iterable_iteration() {
    let out = run_js(
        r#"
console.log(Array.from("ab\u0301").join("|"));
console.log([...new Set("abb")].join("|"));
"#,
    );
    assert_eq!(out, vec!["a|b|́", "a|b"]);
}

#[test]
fn string_repeat_zero_and_empty() {
    let out = run_js(
        r#"
console.log("x".repeat(0) === "");
console.log("".repeat(3));
try {
  console.log("x".repeat(-1));
} catch (e) {
  console.log("RangeError");
}
"#,
    );
    assert_eq!(out, vec!["true", "", "RangeError"]);
}

#[test]
fn string_code_points() {
    let out = run_js(
        r#"
console.log("hello".charCodeAt(1));
console.log("hello".charAt(-1));
console.log(String.fromCharCode(72, 105, 33));
"#,
    );
    assert_eq!(out, vec!["101", "", "Hi!"]);
}

#[test]
fn string_search_methods_with_positions() {
    let out = run_js(
        r#"
console.log("hello world".indexOf("o", 5));
console.log("hello world".search(/world/));
console.log("hello world".match(/o/g).length);
"#,
    );
    assert_eq!(out, vec!["7", "6", "2"]);
}

#[test]
fn string_search_with_offset_positions() {
    let out = run_js(
        r#"
console.log("banana".startsWith("ana", 1));
console.log("banana".startsWith("ana", 2));
console.log("banana".endsWith("na", 4));
console.log("banana".endsWith("na", 5));
"#,
    );
    assert_eq!(out, vec!["true", "false", "true", "false"]);
}

#[test]
fn raw_template_string_keeps_backslashes() {
    let out = run_js(
        r#"
const s = String.raw`line1\nline2`;
console.log(s === "line1\\nline2");
console.log(s.includes("\\n"));
"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

// ── Numeric literals ───────────────────────────────────────

#[test]
fn hex_literal() {
    let out = run_js("console.log(0xFF);");
    assert_eq!(out, vec!["255"]);
}

#[test]
fn octal_literal() {
    let out = run_js("console.log(0o77);");
    assert_eq!(out, vec!["63"]);
}

#[test]
fn binary_literal() {
    let out = run_js("console.log(0b1010);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn scientific_notation() {
    let out = run_js("console.log(1.5e3);");
    assert_eq!(out, vec!["1500"]);
}

#[test]
fn negative_exponent() {
    let out = run_js("console.log(1.5e-1);");
    assert_eq!(out, vec!["0.15"]);
}

// ── JSON ───────────────────────────────────────────────────

#[test]
fn json_parse() {
    let out = run_js(
        r#"
const obj = JSON.parse('{"name":"Alice","age":30}');
console.log(obj.name);
console.log(obj.age);
"#,
    );
    assert_eq!(out, vec!["Alice", "30"]);
}

#[test]
fn json_stringify() {
    let out = run_js(
        r#"
const s = JSON.stringify({ x: 1, y: 2 });
console.log(s);
"#,
    );
    assert!(out[0].contains("\"x\"") && out[0].contains("1"));
}

#[test]
fn json_roundtrip() {
    let out = run_js(
        r#"
const orig = { a: 1, b: "hello", c: true };
const s = JSON.stringify(orig);
const parsed = JSON.parse(s);
console.log(parsed.a);
console.log(parsed.b);
console.log(parsed.c);
"#,
    );
    assert_eq!(out, vec!["1", "hello", "true"]);
}

#[test]
fn test_string_replace_all() {
    let out = run_js(r#"console.log("foo-bar-foo".replaceAll("foo", "baz"));"#);
    assert_eq!(out, vec!["baz-bar-baz"]);
}

#[test]
fn test_string_code_point_at() {
    let out = run_js(r#"console.log("A".codePointAt(0));"#);
    assert_eq!(out, vec!["65"]);
}

