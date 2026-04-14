/// JavaScript async patterns: async generators, for await...of,
/// Promise.any, Promise.race edge cases, async IIFE, error propagation
/// in async, sequential vs concurrent async.

use super::helpers::run_js;

// ===================================================================
// ASYNC GENERATORS
// ===================================================================

#[test] fn async_generator_basic() {
    assert_eq!(run_js(r#"
async function* asyncRange(start, end) {
    for (let i = start; i <= end; i++) {
        yield i;
    }
}
async function main() {
    for await (let n of asyncRange(1, 5)) {
        console.log(n);
    }
}
main();
"#), &["1", "2", "3", "4", "5"]);
}

#[test] fn async_generator_with_await() {
    assert_eq!(run_js(r#"
async function* fetchItems() {
    yield await Promise.resolve("item1");
    yield await Promise.resolve("item2");
    yield await Promise.resolve("item3");
}
async function main() {
    for await (let item of fetchItems()) {
        console.log(item);
    }
}
main();
"#), &["item1", "item2", "item3"]);
}

#[test] fn async_generator_next() {
    assert_eq!(run_js(r#"
async function* gen() {
    yield 10;
    yield 20;
}
async function main() {
    let g = gen();
    let r1 = await g.next();
    console.log(r1.value);
    console.log(r1.done);
    let r2 = await g.next();
    console.log(r2.value);
    let r3 = await g.next();
    console.log(r3.done);
}
main();
"#), &["10", "false", "20", "true"]);
}

// ===================================================================
// FOR AWAIT...OF
// ===================================================================

#[test] fn for_await_of_array() {
    assert_eq!(run_js(r#"
async function main() {
    let promises = [
        Promise.resolve("a"),
        Promise.resolve("b"),
        Promise.resolve("c")
    ];
    for await (let val of promises) {
        console.log(val);
    }
}
main();
"#), &["a", "b", "c"]);
}

// ===================================================================
// PROMISE.ANY
// ===================================================================

#[test] fn promise_any_first_resolved() {
    assert_eq!(run_js(r#"
Promise.any([
    Promise.reject("err1"),
    Promise.resolve("ok"),
    Promise.reject("err2")
]).then(v => console.log(v));
"#), &["ok"]);
}

#[test] fn promise_any_all_reject() {
    assert_eq!(run_js(r#"
Promise.any([
    Promise.reject("e1"),
    Promise.reject("e2"),
    Promise.reject("e3")
]).catch(e => console.log(e instanceof AggregateError));
"#), &["true"]);
}

// ===================================================================
// PROMISE CHAINING ADVANCED
// ===================================================================

#[test] fn promise_then_returns_promise() {
    assert_eq!(run_js(r#"
Promise.resolve(1)
    .then(v => Promise.resolve(v + 1))
    .then(v => Promise.resolve(v * 3))
    .then(v => console.log(v));
"#), &["6"]);
}

#[test] fn promise_chain_error_recovery() {
    assert_eq!(run_js(r#"
Promise.resolve(1)
    .then(v => { throw new Error("oops"); })
    .catch(e => 42)
    .then(v => console.log(v));
"#), &["42"]);
}

#[test] fn promise_chain_finally_passthrough() {
    assert_eq!(run_js(r#"
Promise.resolve("hello")
    .finally(() => console.log("cleanup"))
    .then(v => console.log(v));
"#), &["cleanup", "hello"]);
}

// ===================================================================
// ASYNC ERROR HANDLING
// ===================================================================

#[test] fn async_throw_caught_by_catch() {
    assert_eq!(run_js(r#"
async function fail() {
    throw new Error("async error");
}
fail().catch(e => console.log(e.message));
"#), &["async error"]);
}

#[test] fn async_try_catch_await() {
    assert_eq!(run_js(r#"
async function riskyOp() {
    return Promise.reject("bad");
}
async function main() {
    try {
        await riskyOp();
    } catch (e) {
        console.log("caught: " + e);
    }
}
main();
"#), &["caught: bad"]);
}

// ===================================================================
// ASYNC IIFE
// ===================================================================

#[test] fn async_iife() {
    assert_eq!(run_js(r#"
(async () => {
    let val = await Promise.resolve(42);
    console.log(val);
})();
"#), &["42"]);
}

// ===================================================================
// SEQUENTIAL VS CONCURRENT
// ===================================================================

#[test] fn async_sequential() {
    assert_eq!(run_js(r#"
async function step(n) { return n * 2; }
async function main() {
    let a = await step(1);
    let b = await step(a);
    let c = await step(b);
    console.log(c);
}
main();
"#), &["8"]);
}

#[test] fn async_concurrent_promise_all() {
    assert_eq!(run_js(r#"
async function double(n) { return n * 2; }
async function main() {
    let [a, b, c] = await Promise.all([double(1), double(2), double(3)]);
    console.log(a + "," + b + "," + c);
}
main();
"#), &["2,4,6"]);
}

// ===================================================================
// ASYNC CLASS METHODS
// ===================================================================

#[test] fn async_class_method() {
    assert_eq!(run_js(r#"
class Api {
    async fetch(id) {
        let result = await Promise.resolve("item_" + id);
        return result;
    }
}
async function main() {
    let api = new Api();
    let r = await api.fetch(42);
    console.log(r);
}
main();
"#), &["item_42"]);
}

#[test] fn async_static_method() {
    assert_eq!(run_js(r#"
class Factory {
    static async create(name) {
        let data = await Promise.resolve({ name });
        return data;
    }
}
async function main() {
    let obj = await Factory.create("test");
    console.log(obj.name);
}
main();
"#), &["test"]);
}
