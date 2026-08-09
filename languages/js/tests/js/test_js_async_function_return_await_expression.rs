use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Async Function Return & Await Expression Mechanics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_async_function_implicit_promise_wrapping() {
    let src = r#"
async function getVal() {
    return 42;
}
getVal().then(v => console.log(v + "|" + (getVal() instanceof Promise)));
"#;
    assert_eq!(run_js(src), vec!["42|true"]);
}

#[test]
fn test_js_async_function_await_promise_resolution() {
    let src = r#"
async function compute() {
    const a = await Promise.resolve(10);
    const b = await Promise.resolve(20);
    return a + b;
}
compute().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_async_function_await_primitive_value() {
    let src = r#"
async function getPrimitive() {
    const val = await 100; // Primitive is wrapped in resolved promise
    return val * 2;
}
getPrimitive().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["200"]);
}

#[test]
fn test_js_async_function_await_thenable_object() {
    let src = r#"
async function getThenable() {
    const val = await {
        then(resolve) { resolve("ThenableResolved"); }
    };
    return val;
}
getThenable().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["ThenableResolved"]);
}

#[test]
fn test_js_async_function_await_rejected_promise_throws() {
    let src = r#"
async function fail() {
    try {
        await Promise.reject("AsyncRejection");
    } catch (err) {
        console.log("Caught: " + err);
    }
}
fail();
"#;
    assert_eq!(run_js(src), vec!["Caught: AsyncRejection"]);
}

#[test]
fn test_js_async_function_return_await_preserves_stack() {
    let src = r#"
async function inner() {
    throw new Error("InnerFail");
}
async function outer() {
    return await inner();
}
outer().catch(err => console.log(err.message));
"#;
    assert_eq!(run_js(src), vec!["InnerFail"]);
}

#[test]
fn test_js_async_function_expression_invocation() {
    let src = r#"
const fn = async function(a, b) {
    return a * b;
};
fn(3, 4).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["12"]);
}

#[test]
fn test_js_async_method_in_object_literal() {
    let src = r#"
const obj = {
    async fetch(id) {
        return `Item_${id}`;
    }
};
obj.fetch(99).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["Item_99"]);
}

#[test]
fn test_js_async_method_in_class_definition() {
    let src = r#"
class DataService {
    async loadData() {
        const val = await Promise.resolve("Loaded");
        return val.toUpperCase();
    }
}
new DataService().loadData().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["LOADED"]);
}

#[test]
fn test_js_async_static_method_in_class() {
    let src = r#"
class Helper {
    static async ping() {
        return "pong";
    }
}
Helper.ping().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["pong"]);
}

#[test]
fn test_js_async_function_sequential_vs_parallel_await() {
    let src = r#"
async function testParallel() {
    const p1 = Promise.resolve("P1");
    const p2 = Promise.resolve("P2");
    const r1 = await p1;
    const r2 = await p2;
    return `${r1}+${r2}`;
}
testParallel().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["P1+P2"]);
}

#[test]
fn test_js_async_function_returning_undefined_default() {
    let src = r#"
async function emptyAsync() {}
emptyAsync().then(res => console.log(res === undefined));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_async_iife_execution() {
    let src = r#"
(async () => {
    const res = await Promise.resolve("IIFE");
    console.log(res);
})();
"#;
    assert_eq!(run_js(src), vec!["IIFE"]);
}

#[test]
fn test_js_async_function_constructor_prototype_check() {
    let src = r#"
async function dummy() {}
const AsyncFunction = Object.getPrototypeOf(dummy).constructor;
console.log(AsyncFunction.name);
"#;
    assert_eq!(run_js(src), vec!["AsyncFunction"]);
}

#[test]
fn test_js_async_function_dynamic_instantiation() {
    let src = r#"
const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;
const fn = new AsyncFunction("a", "b", "return (await a) + (await b);");
fn(Promise.resolve(5), Promise.resolve(10)).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["15"]);
}

#[test]
fn test_js_async_function_destructured_parameters() {
    let src = r#"
async function processUser({ name, age }) {
    await Promise.resolve();
    return `${name}:${age}`;
}
processUser({ name: "Bob", age: 25 }).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["Bob:25"]);
}

#[test]
fn test_js_async_function_rest_parameters() {
    let src = r#"
async function sumAll(...numbers) {
    let total = 0;
    for (const n of numbers) {
        total += await Promise.resolve(n);
    }
    return total;
}
sumAll(1, 2, 3, 4).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_async_function_default_parameter_evaluation() {
    let src = r#"
async function fetchConfig(timeout = 500) {
    const t = await Promise.resolve(timeout);
    return t;
}
fetchConfig().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["500"]);
}

#[test]
fn test_js_async_function_await_null_and_undefined() {
    let src = r#"
async function testNullUndef() {
    const v1 = await null;
    const v2 = await undefined;
    return (v1 === null) + "|" + (v2 === undefined);
}
testNullUndef().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_async_function_await_precedence_operations() {
    let src = r#"
async function testPrecedence() {
    const val = await Promise.resolve(10) + 5; // (await Promise.resolve(10)) + 5 = 15
    return val;
}
testPrecedence().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["15"]);
}

#[test]
fn test_js_async_function_return_promise_unwraps() {
    let src = r#"
async function getNested() {
    return Promise.resolve(99);
}
getNested().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["99"]);
}
