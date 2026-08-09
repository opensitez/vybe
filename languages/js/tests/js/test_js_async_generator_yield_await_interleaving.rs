use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Async Generators (`async function*`, `yield`, `await`, `yield*`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_async_generator_basic_next_promise() {
    let src = r#"
async function* asyncGen() {
    yield 10;
    yield 20;
}
(async () => {
    const ag = asyncGen();
    const p1 = await ag.next();
    const p2 = await ag.next();
    console.log(`${p1.value}:${p1.done} | ${p2.value}:${p2.done}`);
})();
"#;
    assert_eq!(run_js(src), vec!["10:false | 20:false"]);
}

#[test]
fn test_js_async_generator_await_and_yield_interleaving() {
    let src = r#"
async function* asyncGen() {
    const a = await Promise.resolve(5);
    yield a * 2;
    const b = await Promise.resolve(10);
    yield a + b;
}
(async () => {
    const res = [];
    for await (const val of asyncGen()) res.push(val);
    console.log(res.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["10,15"]);
}

#[test]
fn test_js_async_generator_yield_star_async_iterable() {
    let src = r#"
async function* inner() {
    yield await Promise.resolve("A");
    yield await Promise.resolve("B");
}
async function* outer() {
    yield* inner();
}
(async () => {
    const items = [];
    for await (const x of outer()) items.push(x);
    console.log(items.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["A,B"]);
}

#[test]
fn test_js_async_generator_yield_star_sync_iterable() {
    let src = r#"
async function* gen() {
    yield* [1, 2, 3];
}
(async () => {
    const items = [];
    for await (const x of gen()) items.push(x);
    console.log(items.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_async_generator_return_method() {
    let src = r#"
async function* gen() {
    yield 1;
    yield 2;
}
(async () => {
    const ag = gen();
    await ag.next();
    const ret = await ag.return("AsyncReturn");
    console.log(`${ret.value}:${ret.done}`);
})();
"#;
    assert_eq!(run_js(src), vec!["AsyncReturn:true"]);
}

#[test]
fn test_js_async_generator_throw_method() {
    let src = r#"
async function* gen() {
    try {
        yield 1;
    } catch (e) {
        yield "HandledInAsyncGen: " + e.message;
    }
}
(async () => {
    const ag = gen();
    await ag.next();
    const res = await ag.throw(new Error("AsyncError"));
    console.log(res.value);
})();
"#;
    assert_eq!(run_js(src), vec!["HandledInAsyncGen: AsyncError"]);
}

#[test]
fn test_js_async_generator_tostringtag() {
    let src = r#"
async function* gen() {}
console.log(gen()[Symbol.toStringTag]);
"#;
    assert_eq!(run_js(src), vec!["AsyncGenerator"]);
}

#[test]
fn test_js_async_generator_symbol_async_iterator_returns_self() {
    let src = r#"
async function* gen() {}
const ag = gen();
console.log(ag[Symbol.asyncIterator]() === ag);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_async_generator_passing_arguments_into_next() {
    let src = r#"
async function* gen() {
    const x = yield "start";
    yield x * 10;
}
(async () => {
    const ag = gen();
    console.log((await ag.next()).value);
    console.log((await ag.next(5)).value);
})();
"#;
    assert_eq!(run_js(src), vec!["start", "50"]);
}

#[test]
fn test_js_async_generator_try_finally_cleanup() {
    let src = r#"
let cleanedUp = false;
async function* gen() {
    try {
        yield 1;
    } finally {
        await Promise.resolve();
        cleanedUp = true;
    }
}
(async () => {
    const ag = gen();
    await ag.next();
    await ag.return();
    console.log(cleanedUp);
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_async_generator_expression_concise_method() {
    let src = r#"
const obj = {
    async *stream() {
        yield await Promise.resolve("S1");
        yield await Promise.resolve("S2");
    }
};
(async () => {
    const res = [];
    for await (const item of obj.stream()) res.push(item);
    console.log(res.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["S1,S2"]);
}

#[test]
fn test_js_async_generator_cannot_be_constructed_with_new() {
    let src = r#"
async function* gen() {}
try {
    new gen();
} catch (e) {
    console.log("AsyncGenerator Constructor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["AsyncGenerator Constructor TypeError"]);
}

#[test]
fn test_js_async_generator_rejection_propagation() {
    let src = r#"
async function* gen() {
    yield 1;
    throw new Error("AsyncGenFailed");
}
(async () => {
    const ag = gen();
    await ag.next();
    try {
        await ag.next();
    } catch (e) {
        console.log("Caught: " + e.message);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["Caught: AsyncGenFailed"]);
}

#[test]
fn test_js_async_generator_queueing_next_calls() {
    let src = r#"
async function* gen() {
    yield 1;
    yield 2;
}
(async () => {
    const ag = gen();
    const p1 = ag.next();
    const p2 = ag.next();
    const [r1, r2] = await Promise.all([p1, p2]);
    console.log(`${r1.value}:${r2.value}`);
})();
"#;
    assert_eq!(run_js(src), vec!["1:2"]);
}

#[test]
fn test_js_async_generator_yield_promise_resolution() {
    let src = r#"
async function* gen() {
    yield Promise.resolve(100); // yield with promise is awaited/unwrapped in iterator result!
}
(async () => {
    const ag = gen();
    const res = await ag.next();
    console.log(res.value);
})();
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_async_generator_yield_star_return_value_capture() {
    let src = r#"
async function* inner() {
    yield 1;
    return "InnerReturned";
}
async function* outer() {
    const ret = yield* inner();
    yield "OuterGot:" + ret;
}
(async () => {
    const items = [];
    for await (const x of outer()) items.push(x);
    console.log(items.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["1,OuterGot:InnerReturned"]);
}

#[test]
fn test_js_async_generator_this_binding() {
    let src = r#"
class AsyncStream {
    #prefix = "Stream";
    async *generate() {
        yield `${this.#prefix}_1`;
    }
}
(async () => {
    const s = new AsyncStream();
    for await (const item of s.generate()) console.log(item);
})();
"#;
    assert_eq!(run_js(src), vec!["Stream_1"]);
}

#[test]
fn test_js_async_generator_empty_yield() {
    let src = r#"
async function* gen() {
    yield;
}
(async () => {
    const ag = gen();
    console.log((await ag.next()).value === undefined);
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_async_generator_exhausted_subsequent_calls() {
    let src = r#"
async function* gen() { yield 1; }
(async () => {
    const ag = gen();
    await ag.next();
    await ag.next();
    const r3 = await ag.next();
    console.log(`${r3.value}:${r3.done}`);
})();
"#;
    assert_eq!(run_js(src), vec!["undefined:true"]);
}

#[test]
fn test_js_async_generator_yield_star_rejected_promise_in_stream() {
    let src = r#"
async function* inner() {
    yield 1;
    throw new Error("StreamError");
}
async function* outer() {
    yield* inner();
}
(async () => {
    try {
        for await (const _ of outer());
    } catch (e) {
        console.log(e.message);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["StreamError"]);
}

#[test]
fn test_js_async_generator_throw_on_completed_generator_rethrows() {
    let src = r#"
async function* gen() { yield 1; }
(async () => {
    const ag = gen();
    await ag.next();
    await ag.next();
    try {
        await ag.throw(new Error("post_done_throw"));
    } catch (e) {
        console.log(e.message);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["post_done_throw"]);
}
