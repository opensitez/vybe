use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Computed Property Names in Destructuring & Binding Patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_computed_property_destructuring_string_variable() {
    let src = r#"
const prop = "foo";
const { [prop]: val } = { foo: "bar" };
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["bar"]);
}

#[test]
fn test_js_computed_property_destructuring_expression_evaluation() {
    let src = r#"
const prefix = "user_";
const { [prefix + "id"]: userId } = { user_id: 1001 };
console.log(userId);
"#;
    assert_eq!(run_js(src), vec!["1001"]);
}

#[test]
fn test_js_computed_property_destructuring_symbol_key() {
    let src = r#"
const sym = Symbol("private");
const { [sym]: val } = { [sym]: "SecretValue" };
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["SecretValue"]);
}

#[test]
fn test_js_computed_property_destructuring_number_key() {
    let src = r#"
const idx = 0;
const { [idx]: first } = { 0: "zero" };
console.log(first);
"#;
    assert_eq!(run_js(src), vec!["zero"]);
}

#[test]
fn test_js_computed_property_destructuring_with_default_value() {
    let src = r#"
const key = "missingKey";
const { [key]: val = "Fallback" } = {};
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["Fallback"]);
}

#[test]
fn test_js_computed_property_destructuring_side_effect_evaluation_order() {
    let src = r#"
let order = [];
function getPropName() {
    order.push("computedKey");
    return "a";
}
const { [getPropName()]: val } = { a: 50 };
console.log(val + "|Order=" + order.join(","));
"#;
    assert_eq!(run_js(src), vec!["50|Order=computedKey"]);
}

#[test]
fn test_js_computed_property_destructuring_alias_same_name_as_key_var() {
    let src = r#"
const k = "targetKey";
const { [k]: targetKey } = { targetKey: "Match" };
console.log(targetKey);
"#;
    assert_eq!(run_js(src), vec!["Match"]);
}

#[test]
fn test_js_computed_property_destructuring_nested_computed_properties() {
    let src = r#"
const outerKey = "outer";
const innerKey = "inner";
const { [outerKey]: { [innerKey]: result } } = { outer: { inner: 42 } };
console.log(result);
"#;
    assert_eq!(run_js(src), vec!["42"]);
}

#[test]
fn test_js_computed_property_destructuring_well_known_symbol() {
    let src = r#"
const obj = { [Symbol.toStringTag]: "CustomModule" };
const { [Symbol.toStringTag]: tag } = obj;
console.log(tag);
"#;
    assert_eq!(run_js(src), vec!["CustomModule"]);
}

#[test]
fn test_js_computed_property_destructuring_assignment_existing_variables() {
    let src = r#"
let val;
const key = "x";
({ [key]: val } = { x: 10 });
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_computed_property_destructuring_array_index_as_string() {
    let src = r#"
const { ["0"]: first, ["1"]: second } = [100, 200];
console.log(`${first}:${second}`);
"#;
    assert_eq!(run_js(src), vec!["100:200"]);
}

#[test]
fn test_js_computed_property_destructuring_boolean_key() {
    let src = r#"
const flag = true;
const { [flag]: val } = { true: "BoolVal" };
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["BoolVal"]);
}

#[test]
fn test_js_computed_property_destructuring_object_implicit_string_conversion() {
    let src = r#"
const keyObj = {
    toString() { return "keyName"; }
};
const { [keyObj]: val } = { keyName: "ValueFromObjKey" };
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["ValueFromObjKey"]);
}

#[test]
fn test_js_computed_property_destructuring_function_parameter() {
    let src = r#"
function extractKey(key, { [key]: value }) {
    return value;
}
console.log(extractKey("score", { score: 95 }));
"#;
    assert_eq!(run_js(src), vec!["95"]);
}

#[test]
fn test_js_computed_property_destructuring_rest_element_combination() {
    let src = r#"
const key = "a";
const { [key]: val, ...rest } = { a: 1, b: 2, c: 3 };
console.log(val + "|" + Object.keys(rest).join(","));
"#;
    assert_eq!(run_js(src), vec!["1|b,c"]);
}

#[test]
fn test_js_computed_property_destructuring_template_literal_key() {
    let src = r#"
const id = 5;
const { [`item_${id}`]: item } = { item_5: "Gadget" };
console.log(item);
"#;
    assert_eq!(run_js(src), vec!["Gadget"]);
}

#[test]
fn test_js_computed_property_destructuring_in_for_of_loop() {
    let src = r#"
const records = [{ k: "a", a: 1 }, { k: "b", b: 2 }];
const results = [];
for (const { k, [k]: val } of records) {
    results.push(val);
}
console.log(results.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_computed_property_destructuring_undefined_computed_key_throws() {
    let src = r#"
let missingVar;
try {
    const { [missingVar]: val } = { undefined: "UndefinedKeyVal" };
    console.log(val); // In JS, undefined is converted to "undefined" key string!
} catch (e) {
    console.log("Error");
}
"#;
    assert_eq!(run_js(src), vec!["UndefinedKeyVal"]);
}

#[test]
fn test_js_computed_property_destructuring_null_computed_key() {
    let src = r#"
const nullKey = null;
const { [nullKey]: val } = { null: "NullKeyVal" };
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["NullKeyVal"]);
}

#[test]
fn test_js_computed_property_destructuring_class_method_param() {
    let src = r#"
class Handler {
    process(key, { [key]: val }) {
        return val * 10;
    }
}
console.log(new Handler().process("amount", { amount: 5 }));
"#;
    assert_eq!(run_js(src), vec!["50"]);
}
