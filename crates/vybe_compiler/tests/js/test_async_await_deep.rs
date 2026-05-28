/// Async/await advanced — for-await, async class methods, async IIFE, error in finally

use super::helpers::run_js;

#[test]
fn async_iife() {
    assert_eq!(run_js(r#"
(async () => {
    const val = await Promise.resolve(42);
    console.log(val);
})();
"#), vec!["42"]);
}

#[test]
fn async_class_method() {
    assert_eq!(run_js(r#"
class DataService {
    async fetch(id) {
        return await Promise.resolve({ id, name: "Item " + id });
    }
}
async function main() {
    const svc = new DataService();
    const item = await svc.fetch(5);
    console.log(item.id);
    console.log(item.name);
}
main();
"#), vec!["5", "Item 5"]);
}

#[test]
fn async_getter_not_directly_supported_workaround() {
    assert_eq!(run_js(r#"
class Loader {
    async load() {
        return await Promise.resolve("data");
    }
}
async function main() {
    const loader = new Loader();
    console.log(await loader.load());
}
main();
"#), vec!["data"]);
}

#[test]
fn async_function_returns_value() {
    assert_eq!(run_js(r#"
async function double(x) {
    return x * 2;
}
async function main() {
    console.log(await double(21));
}
main();
"#), vec!["42"]);
}

#[test]
fn await_inside_loop() {
    assert_eq!(run_js(r#"
async function main() {
    const results = [];
    for (let i = 1; i <= 3; i++) {
        const v = await Promise.resolve(i * i);
        results.push(v);
    }
    console.log(results.join(","));
}
main();
"#), vec!["1,4,9"]);
}

#[test]
fn async_try_catch() {
    assert_eq!(run_js(r#"
async function unsafe() {
    throw new Error("oops");
}
async function main() {
    let caught = null;
    try {
        await unsafe();
    } catch (e) {
        caught = e.message;
    }
    console.log(caught);
}
main();
"#), vec!["oops"]);
}

#[test]
fn async_finally_always_runs() {
    assert_eq!(run_js(r#"
const log = [];
async function main() {
    try {
        await Promise.resolve();
        log.push("try");
    } finally {
        log.push("finally");
    }
    console.log(log.join(","));
}
main();
"#), vec!["try,finally"]);
}

#[test]
fn async_parallel_execution() {
    assert_eq!(run_js(r#"
async function delay(ms, val) {
    return new Promise(resolve => setTimeout(() => resolve(val), ms));
}
async function main() {
    const start = Date.now();
    const [a, b] = await Promise.all([
        delay(10, 1),
        delay(10, 2),
    ]);
    console.log(a + b);
}
main();
"#), vec!["3"]);
}

#[test]
fn async_for_of_with_generator() {
    assert_eq!(run_js(r#"
async function* numbers() {
    yield 1;
    yield 2;
    yield 3;
}
async function main() {
    const sum = [];
    for await (const n of numbers()) {
        sum.push(n);
    }
    console.log(sum.join(","));
}
main();
"#), vec!["1,2,3"]);
}

#[test]
fn async_sequential_accumulator() {
    assert_eq!(run_js(r#"
async function accumulate(fns) {
    let result = 0;
    for (const fn of fns) {
        result = await fn(result);
    }
    return result;
}
async function main() {
    const result = await accumulate([
        async x => x + 1,
        async x => x * 2,
        async x => x + 10,
    ]);
    console.log(result); // ((0+1)*2)+10 = 12
}
main();
"#), vec!["12"]);
}

#[test]
fn async_timeout_pattern() {
    assert_eq!(run_js(r#"
function withTimeout(promise, ms) {
    const timeout = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("timeout")), ms)
    );
    return Promise.race([promise, timeout]);
}
async function main() {
    // fast operation wins
    const fast = Promise.resolve("done");
    const result = await withTimeout(fast, 1000);
    console.log(result);
    // slow would timeout, but we test success path only
}
main();
"#), vec!["done"]);
}
