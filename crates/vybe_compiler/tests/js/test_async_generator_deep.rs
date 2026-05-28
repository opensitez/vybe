/// Async generator patterns — yield with await, for-await-of, return/throw

use super::helpers::run_js;

#[test]
fn async_generator_yields_promises() {
    assert_eq!(run_js(r#"
async function* gen() {
    yield 1;
    yield 2;
    yield 3;
}
async function main() {
    const results = [];
    for await (const v of gen()) results.push(v);
    console.log(results.join(","));
}
main();
"#), vec!["1,2,3"]);
}

#[test]
fn async_generator_with_await() {
    assert_eq!(run_js(r#"
async function* gen() {
    const a = await Promise.resolve(10);
    yield a;
    const b = await Promise.resolve(20);
    yield b;
}
async function main() {
    const results = [];
    for await (const v of gen()) results.push(v);
    console.log(results.join(","));
}
main();
"#), vec!["10,20"]);
}

#[test]
fn async_generator_return_value() {
    assert_eq!(run_js(r#"
async function* gen() {
    yield 1;
    return "done";
}
async function main() {
    const it = gen();
    const r1 = await it.next();
    const r2 = await it.next();
    console.log(r1.value);
    console.log(r2.value);
    console.log(r2.done);
}
main();
"#), vec!["1", "done", "true"]);
}

#[test]
fn async_generator_throw_propagates() {
    assert_eq!(run_js(r#"
async function* gen() {
    try {
        yield 1;
        yield 2;
    } catch (e) {
        yield "caught: " + e;
    }
}
async function main() {
    const it = gen();
    await it.next();
    const r = await it.throw("oops");
    console.log(r.value);
}
main();
"#), vec!["caught: oops"]);
}

#[test]
fn for_await_of_promises_array() {
    assert_eq!(run_js(r#"
async function main() {
    const promises = [
        Promise.resolve(1),
        Promise.resolve(2),
        Promise.resolve(3),
    ];
    const results = [];
    for await (const v of promises) results.push(v);
    console.log(results.join(","));
}
main();
"#), vec!["1,2,3"]);
}

#[test]
fn async_generator_early_return() {
    assert_eq!(run_js(r#"
async function* gen() {
    try {
        yield 1;
        yield 2;
        yield 3;
    } finally {
        console.log("cleanup");
    }
}
async function main() {
    const it = gen();
    await it.next();
    await it.return("stop");
    console.log("done");
}
main();
"#), vec!["cleanup", "done"]);
}

#[test]
fn async_generator_yields_sequential() {
    assert_eq!(run_js(r#"
const order = [];
async function* gen() {
    order.push("before 1");
    yield 1;
    order.push("before 2");
    yield 2;
}
async function main() {
    const it = gen();
    await it.next();
    order.push("after first next");
    await it.next();
    console.log(order.join(","));
}
main();
"#), vec!["before 1,after first next,before 2"]);
}

#[test]
fn async_generator_yield_star() {
    assert_eq!(run_js(r#"
async function* inner() {
    yield "a";
    yield "b";
}
async function* outer() {
    yield* inner();
    yield "c";
}
async function main() {
    const results = [];
    for await (const v of outer()) results.push(v);
    console.log(results.join(","));
}
main();
"#), vec!["a,b,c"]);
}

#[test]
fn async_generator_empty() {
    assert_eq!(run_js(r#"
async function* empty() {}
async function main() {
    const results = [];
    for await (const v of empty()) results.push(v);
    console.log(results.length);
}
main();
"#), vec!["0"]);
}

#[test]
fn async_generator_lazy_evaluation() {
    assert_eq!(run_js(r#"
let count = 0;
async function* gen() {
    while (true) {
        count++;
        yield count;
    }
}
async function main() {
    const it = gen();
    await it.next();
    await it.next();
    await it.next();
    console.log(count);
}
main();
"#), vec!["3"]);
}
