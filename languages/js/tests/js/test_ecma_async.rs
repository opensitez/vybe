use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Async — async/await, Promises
// ═══════════════════════════════════════════════════════════

#[test]
fn async_function_returns_value() {
    let out = run_js(
        r#"
async function getValue() {
    return 42;
}
const p = getValue();
console.log(p);
"#,
    );
    assert!(!out.is_empty());
}

#[test]
fn async_arrow_function() {
    let out = run_js(
        r#"
const fn = async () => "hello";
const p = fn();
console.log(p);
"#,
    );
    assert!(!out.is_empty());
}

#[test]
fn async_function_declaration() {
    let out = run_js(
        r#"
async function process(x) {
    return x * 2;
}
const r = process(21);
console.log(r);
"#,
    );
    assert!(!out.is_empty());
}

#[test]
fn async_class_method() {
    let out = run_js(
        r#"
class Service {
    async fetch(url) {
        return "data from " + url;
    }
}
const s = new Service();
const r = s.fetch("api/users");
console.log(r);
"#,
    );
    assert!(!out.is_empty());
}

#[test]
fn await_expression() {
    let out = run_js(
        r#"
async function main() {
    const x = await 42;
    console.log(x);
}
main();
"#,
    );
    // await on a non-promise should resolve immediately
    assert_eq!(out, vec!["42"]);
}

#[test]
fn async_function_runs_synchronously_until_first_await() {
    let out = run_js(
        r#"
async function demo() {
    console.log("start");
    await 1;
    console.log("end");
}
console.log("before");
demo();
console.log("after");
"#,
    );
    // §27.7.5.3 Await: the continuation after `await` is a microtask —
    // the caller's synchronous "after" prints before "end" (node-verified).
    assert_eq!(out, vec!["before", "start", "after", "end"]);
}

#[test]
fn async_method_can_use_this_before_await() {
    let out = run_js(
        r#"
class Counter {
    constructor() { this.value = 2; }
    async double() {
        console.log(this.value);
        return this.value * 2;
    }
}
const c = new Counter();
console.log(c.double());
"#,
    );
    assert_eq!(out[0], "2");
}

#[test]
fn await_preserves_expression_result() {
    let out = run_js(
        r#"
async function calc() {
    const left = await 5;
    const right = await 7;
    console.log(left + right);
}
calc();
"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn async_arrow_with_parameter() {
    let out = run_js(
        r#"
const double = async x => x * 2;
console.log(double(6));
"#,
    );
    assert!(!out.is_empty());
}

#[test]
fn await_inside_loop() {
    let out = run_js(
        r#"
async function main() {
    for (const value of [1, 2, 3]) {
        console.log(await value);
    }
}
main();
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn async_function_throw_returns_rejected_promise() {
    let out = run_js(
        r#"
async function err() {
    throw "boom";
}
err().catch(e => console.log(e));
"#,
    );
    assert_eq!(out, vec!["boom"]);
}

