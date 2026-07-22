use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Function `length`, `name` Introspection Properties & Descriptors
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_function_length_counts_parameters_before_default() {
    let src = r#"
function fn1(a, b, c) {}
function fn2(a, b = 1, c) {}
function fn3(a = 1, b, c) {}
console.log(`${fn1.length}:${fn2.length}:${fn3.length}`);
"#;
    assert_eq!(run_js(src), vec!["3:1:0"]);
}

#[test]
fn test_js_function_length_excludes_rest_parameter() {
    let src = r#"
function fn(a, b, ...rest) {}
console.log(fn.length);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_function_length_property_descriptor() {
    let src = r#"
function fn() {}
const desc = Object.getOwnPropertyDescriptor(fn, "length");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:true"]); // Configurable is true, writable is false, enumerable is false!
}

#[test]
fn test_js_function_name_property_descriptor() {
    let src = r#"
function fn() {}
const desc = Object.getOwnPropertyDescriptor(fn, "name");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}`);
"#;
    assert_eq!(run_js(src), vec!["false:false:true"]);
}

#[test]
fn test_js_function_name_anonymous_function_variable_inference() {
    let src = r#"
const myFunc = function() {};
console.log(myFunc.name);
"#;
    assert_eq!(run_js(src), vec!["myFunc"]);
}

#[test]
fn test_js_function_name_object_method_shorthand() {
    let src = r#"
const obj = {
    method() {},
    get getter() {},
    set setter(v) {}
};
const getDesc = Object.getOwnPropertyDescriptor(obj, "getter");
const setDesc = Object.getOwnPropertyDescriptor(obj, "setter");

console.log(`${obj.method.name}:${getDesc.get.name}:${setDesc.set.name}`);
"#;
    assert_eq!(run_js(src), vec!["method:get getter:set setter"]);
}

#[test]
fn test_js_function_name_symbol_computed_properties() {
    let src = r#"
const sym = Symbol("mySym");
const obj = {
    [sym]() {}
};
console.log(obj[sym].name);
"#;
    assert_eq!(run_js(src), vec!["[mySym]"]);
}

#[test]
fn test_js_function_name_symbol_without_description() {
    let src = r#"
const sym = Symbol();
const obj = {
    [sym]() {}
};
console.log(obj[sym].name);
"#;
    assert_eq!(run_js(src), vec![""]);
}

#[test]
fn test_js_function_name_redefinition_via_define_property() {
    let src = r#"
function fn() {}
Object.defineProperty(fn, "name", { value: "customName", configurable: true });
console.log(fn.name);
"#;
    assert_eq!(run_js(src), vec!["customName"]);
}

#[test]
fn test_js_function_length_redefinition_via_define_property() {
    let src = r#"
function fn() {}
Object.defineProperty(fn, "length", { value: 99, configurable: true });
console.log(fn.length);
"#;
    assert_eq!(run_js(src), vec!["99"]);
}

#[test]
fn test_js_function_name_class_declaration_and_constructor() {
    let src = r#"
class Person {
    constructor() {}
}
console.log(Person.name + "|" + Person.prototype.constructor.name);
"#;
    assert_eq!(run_js(src), vec!["Person|Person"]);
}

#[test]
fn test_js_function_name_class_expression_variable_inference() {
    let src = r#"
const CustomClass = class {};
console.log(CustomClass.name);
"#;
    assert_eq!(run_js(src), vec!["CustomClass"]);
}

#[test]
fn test_js_function_name_class_expression_named() {
    let src = r#"
const CustomClass = class InternalName {};
console.log(CustomClass.name);
"#;
    assert_eq!(run_js(src), vec!["InternalName"]);
}

#[test]
fn test_js_function_name_bound_function_prefix() {
    let src = r#"
function calc() {}
const boundCalc = calc.bind(null);
const doubleBound = boundCalc.bind(null);
console.log(boundCalc.name + "|" + doubleBound.name);
"#;
    assert_eq!(run_js(src), vec!["bound calc|bound bound calc"]);
}

#[test]
fn test_js_function_prototype_property_descriptor() {
    let src = r#"
function fn() {}
const desc = Object.getOwnPropertyDescriptor(fn, "prototype");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:false"]); // Prototype property is writable, non-enumerable, non-configurable!
}

#[test]
fn test_js_arrow_function_has_no_own_prototype_property() {
    let src = r#"
const arrow = () => {};
console.log(Object.getOwnPropertyDescriptor(arrow, "prototype") === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_function_tostring_preserves_source_formatting() {
    let src = r#"
function mySourceFn(x) { return x + 1; }
console.log(mySourceFn.toString().includes("return x + 1;"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_function_tostring_native_function_formatting() {
    let src = r#"
console.log(Math.sin.toString().includes("[native code]"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_function_tostring_bound_function_formatting() {
    let src = r#"
function orig() {}
const bound = orig.bind(null);
console.log(bound.toString().includes("[native code]"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_function_caller_and_arguments_properties_in_strict_mode_throw() {
    let src = r#"
function strictFn() {
    "use strict";
    try {
        strictFn.caller;
    } catch (e) {
        console.log("Strict Caller TypeError");
    }
}
strictFn();
"#;
    assert_eq!(run_js(src), vec!["Strict Caller TypeError"]);
}
