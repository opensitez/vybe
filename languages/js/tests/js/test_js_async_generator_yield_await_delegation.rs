use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Async Generators, yield, await & yield* Delegation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_async_generator_basic_yield_next() {
    let src = r#"
async function* generate() {
    yield 1;
    yield 2;
    yield 3;
}
(async () => {
    const gen = generate();
    const r1 = await gen.next();
    const r2 = await gen.next();
    const r3 = await gen.next();
    const r4 = await gen.next();
    console.log(`${r1.value},${r2.value},${r3.value}|done=${r4.done}`);
})();
"#;
    assert_eq!(run_js(src), vec!["1,2,3|done=true"]);
}

#[test]
fn test_js_async_generator_await_expression_before_yield() {
    let src = r#"
async function* asyncData() {
    const a = await Promise.resolve(10);
    yield a * 2;
    const b = await Promise.resolve(20);
    yield b * 2;
}
(async () => {
    const gen = asyncData();
    const v1 = (await gen.next()).value;
    const v2 = (await gen.next()).value;
    console.log(`${v1},${v2}`);
})();
"#;
    assert_eq!(run_js(src), vec!["20,40"]);
}

#[test]
fn test_js_async_generator_yield_star_async_iterable_delegation() {
    let src = r#"
async function* subGen() {
    yield "Sub1";
    yield "Sub2";
}
async function* mainGen() {
    yield "Start";
    yield* subGen();
    yield "End";
}
(async () => {
    const results = [];
    for await (const item of mainGen()) {
        results.push(item);
    }
    console.log(results.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["Start,Sub1,Sub2,End"]);
}

#[test]
fn test_js_async_generator_yield_star_sync_iterable_delegation() {
    let src = r#"
async function* delegateSync() {
    yield* [10, 20, 30];
}
(async () => {
    const results = [];
    for await (const n of delegateSync()) {
        results.push(n * 2);
    }
    console.log(results.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["20,40,60"]);
}

#[test]
fn test_js_async_generator_next_with_value_passed_to_yield() {
    let src = r#"
async function* echo() {
    const first = yield "Ready";
    yield `Echo: ${first}`;
}
(async () => {
    const gen = echo();
    await gen.next(); // Start generator
    const r2 = await gen.next("Hello");
    console.log(r2.value);
})();
"#;
    assert_eq!(run_js(src), vec!["Echo: Hello"]);
}

#[test]
fn test_js_async_generator_return_method_early_completion() {
    let src = r#"
async function* counter() {
    try {
        yield 1;
        yield 2;
        yield 3;
    } finally {
        console.log("Async Generator Cleanup");
    }
}
(async () => {
    const gen = counter();
    console.log((await gen.next()).value);
    const r2 = await gen.return("EarlyReturn");
    console.log(`${r2.value}|done=${r2.done}`);
})();
"#;
    assert_eq!(
        run_js(src),
        vec!["1", "Async Generator Cleanup", "EarlyReturn|done=true"]
    );
}

#[test]
fn test_js_async_generator_throw_method_injection() {
    let src = r#"
async function* errorHandler() {
    try {
        yield "Initial";
    } catch (e) {
        yield "Handled: " + e;
    }
}
(async () => {
    const gen = errorHandler();
    await gen.next();
    const r2 = await gen.throw("InjectedError");
    console.log(r2.value);
})();
"#;
    assert_eq!(run_js(src), vec!["Handled: InjectedError"]);
}

#[test]
fn test_js_async_generator_method_in_object() {
    let src = r#"
const obj = {
    async *stream() {
        yield "A";
        yield "B";
    }
};
(async () => {
    const items = [];
    for await (const x of obj.stream()) items.push(x);
    console.log(items.join(""));
})();
"#;
    assert_eq!(run_js(src), vec!["AB"]);
}

#[test]
fn test_js_async_generator_method_in_class() {
    let src = r#"
class Streamer {
    async *items() {
        yield 100;
    }
}
(async () => {
    const gen = new Streamer().items();
    console.log((await gen.next()).value);
})();
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_async_generator_prototype_check() {
    let src = r#"
async function* dummy() {}
const AsyncGeneratorFunction = Object.getPrototypeOf(dummy).constructor;
console.log(AsyncGeneratorFunction.name);
"#;
    assert_eq!(run_js(src), vec!["AsyncGeneratorFunction"]);
}

#[test]
fn test_js_async_generator_symbol_async_iterator_self() {
    let src = r#"
async function* gen() { yield 1; }
const instance = gen();
console.log(instance[Symbol.asyncIterator]() === instance);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_async_generator_yield_promise_resolution() {
    let src = r#"
async function* yieldPromises() {
    yield Promise.resolve("UnwrappedValue");
}
(async () => {
    const gen = yieldPromises();
    const res = await gen.next();
    console.log(res.value);
})();
"#;
    assert_eq!(run_js(src), vec!["UnwrappedValue"]);
}

#[test]
fn test_js_async_generator_yield_star_return_value_capture() {
    let src = r#"
async function* inner() {
    yield 1;
    return "ReturnValue";
}
async function* outer() {
    const ret = yield* inner();
    yield "Captured: " + ret;
}
(async () => {
    const items = [];
    for await (const x of outer()) items.push(x);
    console.log(items.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["1,Captured: ReturnValue"]);
}

#[test]
fn test_js_async_generator_exception_in_body_rejects_next_promise() {
    let src = r#"
async function* thrower() {
    yield 1;
    throw new Error("GeneratorBodyError");
}
(async () => {
    const gen = thrower();
    await gen.next();
    try {
        await gen.next();
    } catch (e) {
        console.log(e.message);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["GeneratorBodyError"]);
}

#[test]
fn test_js_async_generator_concurrent_next_queue_ordering() {
    let src = r#"
async function* sequence() {
    yield 1;
    yield 2;
}
const gen = sequence();
// Calling next concurrently queues promises in FIFO order!
const p1 = gen.next();
const p2 = gen.next();
Promise.all([p1, p2]).then(results => {
    console.log(`${results[0].value},${results[1].value}`);
});
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_async_generator_dynamic_instantiation() {
    let src = r#"
const AsyncGenFn = Object.getPrototypeOf(async function*(){}).constructor;
const genFn = new AsyncGenFn("yield await Promise.resolve('DynamicAsyncGen');");
(async () => {
    const gen = genFn();
    console.log((await gen.next()).value);
})();
"#;
    assert_eq!(run_js(src), vec!["DynamicAsyncGen"]);
}

#[test]
fn test_js_async_generator_yield_star_rejection_propagation() {
    let src = r#"
async function* failingInner() {
    yield "A";
    throw new Error("DelegatedError");
}
async function* outer() {
    yield* failingInner();
}
(async () => {
    const gen = outer();
    await gen.next();
    try {
        await gen.next();
    } catch (e) {
        console.log(e.message);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["DelegatedError"]);
}

#[test]
fn test_js_async_generator_infinite_stream_take() {
    let src = r#"
async function* infiniteNumbers() {
    let i = 1;
    while (true) {
        yield i++;
    }
}
(async () => {
    const results = [];
    for await (const n of infiniteNumbers()) {
        results.push(n);
        if (n === 3) break;
    }
    console.log(results.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_async_generator_empty_yield() {
    let src = r#"
async function* emptyYield() {
    yield;
}
(async () => {
    const res = await emptyYield().next();
    console.log(res.value === undefined);
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_async_generator_yield_star_null_throws_typeerror() {
    let src = r#"
async function* badDelegate() {
    yield* null;
}
(async () => {
    try {
        await badDelegate().next();
    } catch (e) {
        console.log("Delegation Non-Iterable TypeError");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["Delegation Non-Iterable TypeError"]);
}
