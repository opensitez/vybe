use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Function` Constructor (`new Function()`, `AsyncFunction`, `GeneratorFunction`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_function_constructor_basic_execution() {
    let src = r#"
const add = new Function("a", "b", "return a + b;");
console.log(add(10, 20));
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_function_constructor_global_scope_closure_isolation() {
    let src = r#"
const localVal = "OuterLocal";
globalThis.globVal = "GlobalVal";
const fn = new Function("return globVal;"); // Function constructor code executes ONLY in global scope!
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["GlobalVal"]);
}

#[test]
fn test_js_function_constructor_cannot_access_local_lexical_variables() {
    let src = r#"
(() => {
    const hiddenVal = 99;
    const fn = new Function("try { return hiddenVal; } catch(e) { return 'ReferenceError'; }");
    console.log(fn());
})();
"#;
    assert_eq!(run_js(src), vec!["ReferenceError"]);
}

#[test]
fn test_js_function_constructor_comma_separated_parameters() {
    let src = r#"
const sum3 = new Function("a, b, c", "return a + b + c;");
console.log(sum3(1, 2, 3));
"#;
    assert_eq!(run_js(src), vec!["6"]);
}

#[test]
fn test_js_function_constructor_no_arguments_empty_body() {
    let src = r#"
const emptyFn = new Function();
console.log(emptyFn() === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_generator_function_constructor() {
    let src = r#"
const GeneratorFunction = Object.getPrototypeOf(function*(){}).constructor;
const gen = new GeneratorFunction("a", "yield a * 10; yield a * 20;");
const g = gen(5);
console.log(`${g.next().value}:${g.next().value}`);
"#;
    assert_eq!(run_js(src), vec!["50:100"]);
}

#[test]
fn test_js_async_function_constructor() {
    let src = r#"
const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;
const asyncAdd = new AsyncFunction("a", "b", "return await Promise.resolve(a + b);");
(async () => {
    console.log(await asyncAdd(15, 25));
})();
"#;
    assert_eq!(run_js(src), vec!["40"]);
}

#[test]
fn test_js_async_generator_function_constructor() {
    let src = r#"
const AsyncGeneratorFunction = Object.getPrototypeOf(async function*(){}).constructor;
const asyncGen = new AsyncGeneratorFunction("a", "yield await Promise.resolve(a * 2);");
(async () => {
    const ag = asyncGen(10);
    console.log((await ag.next()).value);
})();
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_function_constructor_destructuring_parameters() {
    let src = r#"
const fn = new Function("{ name, age }", "return name + ' is ' + age;");
console.log(fn({ name: "Bob", age: 30 }));
"#;
    assert_eq!(run_js(src), vec!["Bob is 30"]);
}

#[test]
fn test_js_function_constructor_default_parameters() {
    let src = r#"
const fn = new Function("x = 10", "y = 20", "return x + y;");
console.log(fn() + "|" + fn(5));
"#;
    assert_eq!(run_js(src), vec!["30|25"]);
}

#[test]
fn test_js_function_constructor_rest_parameters() {
    let src = r#"
const sumRest = new Function("...nums", "return nums.reduce((a, b) => a + b, 0);");
console.log(sumRest(1, 2, 3, 4));
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_function_constructor_name_is_anonymous() {
    let src = r#"
const fn = new Function("return 1;");
console.log(fn.name);
"#;
    assert_eq!(run_js(src), vec!["anonymous"]);
}

#[test]
fn test_js_function_constructor_prototype_chain() {
    let src = r#"
const fn = new Function();
console.log(Object.getPrototypeOf(fn) === Function.prototype);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_function_constructor_without_new_operator() {
    let src = r#"
const fn = Function("a", "return a * 2;"); // Function(...) without 'new' returns a new function object identically!
console.log(fn(10));
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_function_constructor_invalid_syntax_throws_syntaxerror() {
    let src = r#"
try {
    new Function("invalid { syntax }");
} catch (e) {
    console.log("Function Constructor SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Function Constructor SyntaxError"]);
}

#[test]
fn test_js_function_constructor_strict_mode_directive() {
    let src = r#"
const fn = new Function("'use strict'; try { delete Object.prototype; } catch(e) { return 'StrictEnforced'; }");
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["StrictEnforced"]);
}

#[test]
fn test_js_function_constructor_tostring_output() {
    let src = r#"
const fn = new Function("a", "b", "return a + b;");
console.log(fn.toString().includes("return a + b;"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_function_constructor_created_function_can_be_used_as_constructor() {
    let src = r#"
const MyClass = new Function("val", "this.val = val;");
const inst = new MyClass("DynamicInst");
console.log(inst.val);
"#;
    assert_eq!(run_js(src), vec!["DynamicInst"]);
}

#[test]
fn test_js_function_constructor_length_property() {
    let src = r#"
const fn1 = new Function("a", "b", "c", "return 0;");
const fn2 = new Function("a, b = 1", "c", "return 0;");
console.log(`${fn1.length}:${fn2.length}`);
"#;
    assert_eq!(run_js(src), vec!["3:1"]);
}

#[test]
fn test_js_function_constructor_eval_comparison() {
    let src = r#"
const fnEval = eval("(function(a) { return a * 3; })");
const fnConst = new Function("a", "return a * 3;");
console.log((fnEval(4) === fnConst(4)) + "|" + (fnEval.name !== fnConst.name));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_function_prototype_parent_is_object_prototype() {
    let src = r#"
console.log(Object.getPrototypeOf(Function.prototype) === Object.prototype);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

