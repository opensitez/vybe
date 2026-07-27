use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Template Literal Expression Interpolation Evaluation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_template_literal_basic_variable_interpolation() {
    let src = r#"
const name = "Alice";
const age = 30;
console.log(`User ${name} is ${age} years old.`);
"#;
    assert_eq!(run_js(src), vec!["User Alice is 30 years old."]);
}

#[test]
fn test_js_template_literal_arithmetic_expressions() {
    let src = r#"
const a = 15, b = 25;
console.log(`Sum: ${a + b}, Product: ${a * b}`);
"#;
    assert_eq!(run_js(src), vec!["Sum: 40, Product: 375"]);
}

#[test]
fn test_js_template_literal_function_call_expressions() {
    let src = r#"
function formatName(n) { return n.toUpperCase(); }
console.log(`Hello ${formatName("bob")}`);
"#;
    assert_eq!(run_js(src), vec!["Hello BOB"]);
}

#[test]
fn test_js_template_literal_ternary_operator_expression() {
    let src = r#"
const isMember = true;
console.log(`Fee: $${isMember ? 10 : 50}`);
"#;
    assert_eq!(run_js(src), vec!["Fee: $10"]);
}

#[test]
fn test_js_template_literal_object_property_access() {
    let src = r#"
const user = { details: { city: "Paris" } };
console.log(`Location: ${user.details.city}`);
"#;
    assert_eq!(run_js(src), vec!["Location: Paris"]);
}

#[test]
fn test_js_template_literal_array_method_expressions() {
    let src = r#"
const items = ["apple", "banana"];
console.log(`Items: ${items.join(" & ")}`);
"#;
    assert_eq!(run_js(src), vec!["Items: apple & banana"]);
}

#[test]
fn test_js_template_literal_nested_template_literals() {
    let src = r#"
const isLoggedIn = true;
const user = "Charlie";
console.log(`Status: ${isLoggedIn ? `Welcome back ${user}` : "Guest"}`);
"#;
    assert_eq!(run_js(src), vec!["Status: Welcome back Charlie"]);
}

#[test]
fn test_js_template_literal_implicit_to_string_conversion() {
    let src = r#"
const customObj = {
    toString() { return "CustomObjString"; }
};
console.log(`Result: ${customObj}`);
"#;
    assert_eq!(run_js(src), vec!["Result: CustomObjString"]);
}

#[test]
fn test_js_template_literal_to_string_throw_bubbles_error() {
    let src = r#"
const bad = {
    toString() { throw new Error("toStringFail"); }
};
try {
    console.log(`X${bad}`);
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["toStringFail"]);
}

#[test]
fn test_js_template_literal_null_and_undefined_coercion() {
    let src = r#"
console.log(`Values: ${null} and ${undefined}`);
"#;
    assert_eq!(run_js(src), vec!["Values: null and undefined"]);
}

#[test]
fn test_js_template_literal_symbol_interpolation_throws_typeerror() {
    let src = r#"
const sym = Symbol("id");
try {
    const str = `Sym: ${sym}`;
} catch (e) {
    console.log("Symbol Implicit String Conversion TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Symbol Implicit String Conversion TypeError"]
    );
}

#[test]
fn test_js_template_literal_backslash_before_dollar_does_not_escape_interpolation() {
    let src = r#"
const value = 99;
console.log(`literal: \${value} | computed: ${value}`);
"#;
    assert_eq!(
        run_js(src),
        vec!["literal: \\99 | computed: 99"]
    );
}

#[test]
fn test_js_template_literal_explicitly_constructed_expression_marker() {
    let src = r#"
const suffix = "value}";
console.log(`literal: ${"${" + suffix}`);
"#;
    assert_eq!(run_js(src), vec!["literal: ${value}"]);
}

#[test]
fn test_js_template_literal_side_effect_expressions_order() {
    let src = r#"
let counter = 0;
function inc() { return ++counter; }
console.log(`${inc()}-${inc()}-${inc()}`);
"#;
    assert_eq!(run_js(src), vec!["1-2-3"]);
}

#[test]
fn test_js_template_literal_regex_literal_expression() {
    let src = r#"
console.log(`Match: ${/abc/.test("abcdef")}`);
"#;
    assert_eq!(run_js(src), vec!["Match: true"]);
}

#[test]
fn test_js_template_literal_iife_interpolation() {
    let src = r#"
console.log(`Computed: ${(() => 5 * 5)()}`);
"#;
    assert_eq!(run_js(src), vec!["Computed: 25"]);
}

#[test]
fn test_js_template_literal_bigint_interpolation() {
    let src = r#"
const big = 9007199254740991n;
console.log(`BigInt: ${big}`);
"#;
    assert_eq!(run_js(src), vec!["BigInt: 9007199254740991"]);
}

#[test]
fn test_js_template_literal_assignment_expression() {
    let src = r#"
let val;
console.log(`Assigned: ${val = 100}`);
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["Assigned: 100", "100"]);
}

#[test]
fn test_js_template_literal_logical_assignment_expression() {
    let src = r#"
let config = { port: 0 };
console.log(`Port: ${config.port ||= 8080}`);
"#;
    assert_eq!(run_js(src), vec!["Port: 8080"]);
}

#[test]
fn test_js_template_literal_optional_chaining_expression() {
    let src = r#"
const data = { user: null };
console.log(`Address: ${data.user?.address?.city}`);
"#;
    assert_eq!(run_js(src), vec!["Address: undefined"]);
}

#[test]
fn test_js_template_literal_nullish_coalescing_expression() {
    let src = r#"
const value = null;
console.log(`Output: ${value ?? "DefaultFallback"}`);
"#;
    assert_eq!(run_js(src), vec!["Output: DefaultFallback"]);
}

#[test]
fn test_js_template_literal_escaped_backticks_and_dollars() {
    let src = r#"
const amount = 50;
console.log(`Price: \$${amount} \`code\``);
"#;
    assert_eq!(run_js(src), vec!["Price: $50 `code`"]);
}

#[test]
fn test_js_template_literal_multi_level_nesting() {
    let src = r#"
const a = 1;
console.log(`Level1: ${`Level2: ${`Level3: ${a}`}`}`);
"#;
    assert_eq!(run_js(src), vec!["Level1: Level2: Level3: 1"]);
}

#[test]
fn test_js_template_literal_multiline_text_with_interpolation() {
    let src = r#"
const n = 3;
console.log(`start
value=${n}
end`);
"#;
    assert_eq!(run_js(src), vec!["start\nvalue=3\nend"]);
}

#[test]
fn test_js_template_literal_expression_side_effects() {
    let src = r#"
const trace = [];
const payload = {
    get value() {
        trace.push("read");
        return 7;
    }
};
console.log(`${payload.value}:${payload.value}`);
console.log(trace.length);
"#;
    assert_eq!(run_js(src), vec!["7:7", "2"]);
}

#[test]
fn test_js_template_literal_expression_error_bubbles() {
    let src = r#"
function fail() {
    throw new Error("boom");
}
try {
    console.log(`A${fail()}B`);
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["boom"]);
}

#[test]
fn test_js_template_literal_string_raw_preserves_backslashes() {
    let src = r#"
const value = String.raw`C:\Program Files\Node`;
console.log(`Path: ${value}`);
"#;
    assert_eq!(run_js(src), vec!["Path: C:\\Program Files\\Node"]);
}

#[test]
fn test_js_template_literal_escaped_dollar_brace_is_not_interpolated() {
    let src = r#"
console.log(`\u0024{ignored} and \u0024{alsoIgnored}`);
"#;
    assert_eq!(run_js(src), vec!["${ignored} and ${alsoIgnored}"]);
}

#[test]
fn test_js_template_literal_substitutions_evaluated_left_to_right_then_error_bubbles() {
    let src = r#"
const trace = [];
function boom() {
    trace.push("boom");
    throw new Error("interpolation-failed");
}
try {
    console.log(`a=${(() => { trace.push("first"); return "1"; })()} b=${boom()} c=${"never"}`);
} catch (e) {
    console.log(e.message);
    console.log(trace.join(","));
}
"#;
    assert_eq!(run_js(src), vec!["interpolation-failed", "first,boom"]);
}

#[test]
fn test_js_template_literal_nested_expression_with_trailing_space_preserves_spacing() {
    let src = r#"
function format(label, value) { return `${label}: ${value}`; }
console.log(`${format("a", 1)}|${format("b", 2)}`);
"#;
    assert_eq!(run_js(src), vec!["a: 1|b: 2"]);
}

#[test]
fn test_js_template_literal_tag_receives_raw_and_cooked_strings() {
    let src = r#"
function capture(strings, value) {
    console.log(strings.length);
    console.log(strings.raw.length);
    console.log(strings[0] === "a\nb");
    console.log(strings.raw[0] === "a\\nb");
    console.log(value);
    console.log(strings[1] === "c");
    console.log(strings.raw[1] === "c");
}
capture`a\nb${42}c`;
"#;
    assert_eq!(
        run_js(src),
        vec!["2", "2", "true", "true", "42", "true", "true"]
    );
}

#[test]
fn test_js_template_literal_tagged_template_strings_preserve_cooked_and_raw() {
    let src = r#"
function capture(strings, value) {
    console.log(strings.length);
    console.log(strings[0] === "a\nb");
    console.log(strings.raw[0] === "a\\nb");
    return value;
}
console.log(capture`a\nb${41 + 1}c`);
"#;
    assert_eq!(
        run_js(src),
        vec!["2", "true", "true", "42"]
    );
}

#[test]
fn test_js_template_literal_typeof_expression_in_interpolation() {
    let src = r#"
console.log(`Type: ${typeof null} / ${typeof [1, 2]} / ${typeof undefined}`);
"#;
    assert_eq!(run_js(src), vec!["Type: object / object / undefined"]);
}

#[test]
fn test_js_template_literal_computed_property_and_binary_expression() {
    let src = r#"
const map = { a: 1, b: 2 };
const keys = ["a", "b"];
console.log(`${map[keys[0]] + map[keys[1]]}`);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_template_literal_expression_evaluation_order_with_comma_operator() {
    let src = r#"
let x = 0;
console.log(`value=${(x++, x += 10)}`);
console.log(x);
"#;
    assert_eq!(run_js(src), vec!["value=11", "11"]);
}

#[test]
fn test_js_template_literal_tagged_arrays_preserve_segment_shapes() {
    let src = r#"
function capture(strings) {
    console.log(strings.length);
    console.log(strings.raw.length);
    console.log(strings[0]);
    console.log(strings[1]);
    console.log(strings[2]);
    console.log(strings.raw[0] === "a\\nb");
}

capture`a\nb${1}x${2}y`;
"#;
    assert_eq!(
        run_js(src),
        vec!["3", "3", "a\nb", "x", "y", "true"]
    );
}
