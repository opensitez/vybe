use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Object Destructuring, Default Values & Property Aliases
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_destructuring_basic_properties() {
    let src = r#"
const { a, b } = { a: 10, b: 20 };
console.log(a + "|" + b);
"#;
    assert_eq!(run_js(src), vec!["10|20"]);
}

#[test]
fn test_js_object_destructuring_property_alias() {
    let src = r#"
const { x: width, y: height } = { x: 1920, y: 1080 };
console.log(width + "x" + height);
"#;
    assert_eq!(run_js(src), vec!["1920x1080"]);
}

#[test]
fn test_js_object_destructuring_default_values() {
    let src = r#"
const { port = 8080, host = "localhost" } = { port: 3000 };
console.log(host + ":" + port);
"#;
    assert_eq!(run_js(src), vec!["localhost:3000"]);
}

#[test]
fn test_js_object_destructuring_alias_with_default_value() {
    let src = r#"
const { max: limit = 100 } = {};
console.log(limit);
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_object_destructuring_null_target_throws_typeerror() {
    let src = r#"
try {
    const { a } = null;
} catch (e) {
    console.log("Destructure Null TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Destructure Null TypeError"]);
}

#[test]
fn test_js_object_destructuring_undefined_target_throws_typeerror() {
    let src = r#"
try {
    const { a } = undefined;
} catch (e) {
    console.log("Destructure Undefined TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Destructure Undefined TypeError"]);
}

#[test]
fn test_js_object_destructuring_undefined_property_triggers_default() {
    let src = r#"
const { val = "Default" } = { val: undefined };
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["Default"]);
}

#[test]
fn test_js_object_destructuring_null_property_does_not_trigger_default() {
    let src = r#"
const { val = "Default" } = { val: null };
console.log(val === null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_destructuring_rest_element() {
    let src = r#"
const { id, ...details } = { id: 1, name: "Alice", role: "admin" };
console.log(id + "|" + Object.keys(details).join(","));
"#;
    assert_eq!(run_js(src), vec!["1|name,role"]);
}

#[test]
fn test_js_object_destructuring_assignment_existing_variables() {
    let src = r#"
let x, y;
({ x, y } = { x: 1, y: 2 });
console.log(x + "|" + y);
"#;
    assert_eq!(run_js(src), vec!["1|2"]);
}

#[test]
fn test_js_object_destructuring_side_effect_in_default_initializer() {
    let src = r#"
let count = 0;
function getDefault() { return ++count; }

const { a = getDefault() } = { a: 50 };
const { b = getDefault() } = {};
console.log(a + "|" + b + "|count=" + count);
"#;
    assert_eq!(run_js(src), vec!["50|1|count=1"]);
}

#[test]
fn test_js_object_destructuring_prototype_properties() {
    let src = r#"
const proto = { inherited: "parent" };
const obj = Object.create(proto);
obj.own = "child";

const { own, inherited } = obj;
console.log(own + "|" + inherited);
"#;
    assert_eq!(run_js(src), vec!["child|parent"]);
}

#[test]
fn test_js_object_destructuring_rest_element_excludes_inherited() {
    let src = r#"
const proto = { inherited: 100 };
const obj = Object.create(proto);
obj.own = 200;

const { ...rest } = obj;
console.log(Object.hasOwn(rest, "own") + "|" + Object.hasOwn(rest, "inherited"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_object_destructuring_symbol_properties() {
    let src = r#"
const sym = Symbol("id");
const { [sym]: idVal } = { [sym]: "SYM-123" };
console.log(idVal);
"#;
    assert_eq!(run_js(src), vec!["SYM-123"]);
}

#[test]
fn test_js_object_destructuring_primitive_coercion() {
    let src = r#"
const { length, toFixed } = 123.45;
console.log(length + "|" + (typeof toFixed));
"#;
    assert_eq!(run_js(src), vec!["undefined|function"]);
}

#[test]
fn test_js_object_destructuring_string_primitive() {
    let src = r#"
const { length, slice } = "hello";
console.log(length + "|" + typeof slice);
"#;
    assert_eq!(run_js(src), vec!["5|function"]);
}

#[test]
fn test_js_object_destructuring_getter_invocation() {
    let src = r#"
let readCount = 0;
const obj = {
    get prop() { readCount++; return "Value"; }
};
const { prop } = obj;
console.log(prop + "|Reads=" + readCount);
"#;
    assert_eq!(run_js(src), vec!["Value|Reads=1"]);
}

#[test]
fn test_js_object_destructuring_assignment_returns_right_hand_side() {
    let src = r#"
let a, b;
const rhs = { a: 1, b: 2 };
const res = ({ a, b } = rhs);
console.log(res === rhs);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_destructuring_empty_pattern() {
    let src = r#"
const {} = { x: 1 };
console.log("Empty Pattern Succeeded");
"#;
    assert_eq!(run_js(src), vec!["Empty Pattern Succeeded"]);
}

#[test]
fn test_js_object_destructuring_var_let_const_scoping() {
    let src = r#"
{
    var { v } = { v: "varVal" };
    let { l } = { l: "letVal" };
}
console.log(v + "|" + (typeof l === "undefined"));
"#;
    assert_eq!(run_js(src), vec!["varVal|true"]);
}
