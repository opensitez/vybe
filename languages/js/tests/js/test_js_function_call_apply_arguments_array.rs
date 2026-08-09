use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Function.prototype.call()`, `apply()` & `arguments` Object
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_function_call_explicit_this_and_arguments() {
    let src = r#"
function greet(greeting, punctuation) {
    return `${greeting} ${this.name}${punctuation}`;
}
const user = { name: "Bob" };
console.log(greet.call(user, "Hello", "!"));
"#;
    assert_eq!(run_js(src), vec!["Hello Bob!"]);
}

#[test]
fn test_js_function_apply_explicit_this_and_array_arguments() {
    let src = r#"
function greet(greeting, punctuation) {
    return `${greeting} ${this.name}${punctuation}`;
}
const user = { name: "Bob" };
console.log(greet.apply(user, ["Hi", "?"]));
"#;
    assert_eq!(run_js(src), vec!["Hi Bob?"]);
}

#[test]
fn test_js_function_apply_array_like_arguments() {
    let src = r#"
function sum(a, b, c) {
    return a + b + c;
}
const arrayLike = { 0: 10, 1: 20, 2: 30, length: 3 };
console.log(sum.apply(null, arrayLike));
"#;
    assert_eq!(run_js(src), vec!["60"]);
}

#[test]
fn test_js_arguments_object_indexing_and_length() {
    let src = r#"
function test() {
    return `${arguments[0]}:${arguments[1]}:${arguments.length}`;
}
console.log(test("x", "y"));
"#;
    assert_eq!(run_js(src), vec!["x:y:2"]);
}

#[test]
fn test_js_arguments_object_aliasing_in_non_strict_mode() {
    let src = r#"
function mutate(a) {
    a = 99; // Mutating parameter updates arguments[0] in non-strict mode!
    return arguments[0];
}
console.log(mutate(10));
"#;
    assert_eq!(run_js(src), vec!["99"]);
}

#[test]
fn test_js_arguments_object_no_aliasing_in_strict_mode() {
    let src = r#"
function mutate(a) {
    "use strict";
    a = 99; // Mutating parameter does NOT update arguments[0] in strict mode!
    return arguments[0];
}
console.log(mutate(10));
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_arguments_callee_in_non_strict() {
    let src = r#"
function testCallee() {
    return arguments.callee === testCallee;
}
console.log(testCallee());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_arguments_callee_in_strict_mode_throws_typeerror() {
    let src = r#"
function testCallee() {
    "use strict";
    try {
        arguments.callee;
    } catch (e) {
        console.log("Strict arguments.callee TypeError");
    }
}
testCallee();
"#;
    assert_eq!(run_js(src), vec!["Strict arguments.callee TypeError"]);
}

#[test]
fn test_js_function_call_primitive_this_boxed_in_non_strict() {
    let src = r#"
function getThisType() {
    return typeof this;
}
console.log(getThisType.call("str_prim"));
"#;
    assert_eq!(run_js(src), vec!["object"]);
}

#[test]
fn test_js_function_call_primitive_this_unboxed_in_strict() {
    let src = r#"
function getThisType() {
    "use strict";
    return typeof this;
}
console.log(getThisType.call("str_prim"));
"#;
    assert_eq!(run_js(src), vec!["string"]);
}

#[test]
fn test_js_function_apply_null_args_treated_as_empty() {
    let src = r#"
function getArgCount() {
    return arguments.length;
}
console.log(getArgCount.apply(null, null) + "|" + getArgCount.apply(null, undefined));
"#;
    assert_eq!(run_js(src), vec!["0|0"]);
}

#[test]
fn test_js_function_apply_non_object_non_null_args_throws_typeerror() {
    let src = r#"
function fn() {}
try {
    fn.apply(null, 12345);
} catch (e) {
    console.log("apply Non-Object Args TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["apply Non-Object Args TypeError"]);
}

#[test]
fn test_js_arguments_object_symbol_iterator() {
    let src = r#"
function test() {
    return [...arguments].join(",");
}
console.log(test(1, 2, 3));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_arguments_object_tostringtag_is_arguments() {
    let src = r#"
function test() {
    return Object.prototype.toString.call(arguments);
}
console.log(test());
"#;
    assert_eq!(run_js(src), vec!["[object Arguments]"]);
}

#[test]
fn test_js_function_call_method_borrowing() {
    let src = r#"
const obj1 = { val: 10 };
const obj2 = { val: 20, getVal() { return this.val; } };
console.log(obj2.getVal.call(obj1));
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_function_apply_math_max_spread() {
    let src = r#"
const numbers = [5, 20, 15];
console.log(Math.max.apply(null, numbers));
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_arguments_object_extra_arguments() {
    let src = r#"
function sum(a, b) {
    let total = a + b;
    for (let i = 2; i < arguments.length; i++) {
        total += arguments[i];
    }
    return total;
}
console.log(sum(1, 2, 3, 4));
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_arguments_object_default_parameter_breaks_aliasing() {
    let src = r#"
function fn(a = 10) {
    a = 99; // Default parameters prevent parameter-arguments aliasing even in non-strict mode!
    return arguments[0];
}
console.log(fn(5));
"#;
    assert_eq!(run_js(src), vec!["5"]);
}

#[test]
fn test_js_arguments_object_rest_parameter_breaks_aliasing() {
    let src = r#"
function fn(a, ...rest) {
    a = 99;
    return arguments[0];
}
console.log(fn(5));
"#;
    assert_eq!(run_js(src), vec!["5"]);
}

#[test]
fn test_js_call_non_function_throws_typeerror() {
    let src = r#"
try {
    Function.prototype.call.call("not_a_function");
} catch (e) {
    console.log("call Non-Function TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["call Non-Function TypeError"]);
}

#[test]
fn test_js_arguments_object_destructuring_parameter_breaks_aliasing() {
    let src = r#"
function fn({ a }) {
    a = 99;
    return arguments[0].a;
}
console.log(fn({ a: 5 }));
"#;
    assert_eq!(run_js(src), vec!["5"]);
}
