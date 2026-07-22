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
    assert_eq!(run_js(src), vec!["HELLO BOB"]);
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
