use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Implicit Type Coercion (`+` Operator Addition vs String Concatenation)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_plus_operator_string_and_number() {
    let src = r#"
console.log([
    "5" + 5,
    5 + "5",
    "5" + 5 + 5,
    5 + 5 + "5"
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["55|55|555|105"]);
}

#[test]
fn test_js_plus_operator_boolean_and_number() {
    let src = r#"
console.log([
    true + 1,
    false + 1,
    true + true,
    false + false
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["2|1|2|0"]);
}

#[test]
fn test_js_plus_operator_boolean_and_string() {
    let src = r#"
console.log([
    "val:" + true,
    false + ":val"
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["val:true|false:val"]);
}

#[test]
fn test_js_plus_operator_null_and_undefined() {
    let src = r#"
console.log([
    null + 5,
    undefined + 5,
    null + "str",
    undefined + "str"
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["5|NaN|nullstr|undefinedstr"]);
}

#[test]
fn test_js_plus_operator_arrays() {
    let src = r#"
console.log([
    [1, 2] + [3, 4],
    [] + [],
    [1] + 2
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["1,23,4||12"]);
}

#[test]
fn test_js_plus_operator_objects() {
    let src = r#"
console.log([
    {} + [],
    [] + {},
    "res:" + {}
].join("|"));
"#;
    assert_eq!(
        run_js(src),
        vec!["[object Object]|[object Object]|res:[object Object]"]
    );
}

#[test]
fn test_js_arithmetic_operators_force_numeric_coercion() {
    let src = r#"
console.log([
    "10" - "2",
    "10" * "2",
    "10" / "2",
    "10" % "3"
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["8|20|5|1"]); // Arithmetic -, *, /, % force numeric conversion unlike +!
}

#[test]
fn test_js_arithmetic_operators_with_boolean_null_undefined() {
    let src = r#"
console.log([
    true - false,
    null - 5,
    undefined - 5
].join("|"));
"#;
    assert_eq!(run_js(src), vec!["1|-5|NaN"]);
}

#[test]
fn test_js_plus_operator_symbol_throws_typeerror() {
    let src = r#"
try {
    Symbol("a") + "b";
} catch (e) {
    console.log("Symbol Concatenation TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Symbol Concatenation TypeError"]);
}

#[test]
fn test_js_plus_operator_bigint_and_string_concatenation() {
    let src = r#"
console.log((10n + "5") + "|" + ("5" + 10n));
"#;
    assert_eq!(run_js(src), vec!["105|510"]); // BigInt + string performs string concatenation!
}

#[test]
fn test_js_plus_operator_bigint_and_number_throws_typeerror() {
    let src = r#"
try {
    eval("10n + 5");
} catch (e) {
    console.log("BigInt Number Addition TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["BigInt Number Addition TypeError"]);
}

#[test]
fn test_js_plus_operator_object_toprimitive_default_hint() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        return hint === "default" ? "defaultVal" : "otherVal";
    }
};
console.log(obj + 10); // + operator uses "default" hint!
"#;
    assert_eq!(run_js(src), vec!["defaultVal10"]);
}

#[test]
fn test_js_plus_operator_object_valueof_returning_number() {
    let src = r#"
const obj1 = { valueOf: () => 10 };
const obj2 = { valueOf: () => 20 };
console.log(obj1 + obj2);
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_plus_operator_object_tostring_returning_string() {
    let src = r#"
const obj = { toString: () => "hello" };
console.log(obj + 10);
"#;
    assert_eq!(run_js(src), vec!["hello10"]);
}

#[test]
fn test_js_subtraction_operator_object_toprimitive_number_hint() {
    let src = r#"
const obj = {
    [Symbol.toPrimitive](hint) {
        return hint === "number" ? 50 : 0;
    }
};
console.log(obj - 10); // - operator uses "number" hint!
"#;
    assert_eq!(run_js(src), vec!["40"]);
}

#[test]
fn test_js_plus_operator_date_object_default_hint_is_string() {
    let src = r#"
const d = new Date(0);
console.log(typeof (d + 10)); // Date default hint is "string"!
"#;
    assert_eq!(run_js(src), vec!["string"]);
}

#[test]
fn test_js_unary_plus_date_object_number_hint() {
    let src = r#"
const d = new Date(0);
console.log(typeof (+d)); // Unary + forces "number" hint!
"#;
    assert_eq!(run_js(src), vec!["number"]);
}

#[test]
fn test_js_implicit_coercion_in_template_literals() {
    let src = r#"
console.log(`${10}${true}${null}${undefined}`);
"#;
    assert_eq!(run_js(src), vec!["10truenullundefined"]);
}

#[test]
fn test_js_implicit_coercion_if_statement_condition() {
    let src = r#"
const check = (val) => val ? "truthy" : "falsy";
console.log(`${check("0")}:${check(0)}:${check([])}`);
"#;
    assert_eq!(run_js(src), vec!["truthy:falsy:truthy"]);
}

#[test]
fn test_js_implicit_coercion_bitwise_operators() {
    let src = r#"
console.log(`${"5" | "3"}:${true << 2}:${false | 1}`);
"#;
    assert_eq!(run_js(src), vec!["7:4:1"]);
}
