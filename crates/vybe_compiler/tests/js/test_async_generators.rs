/// Async generators — async function*, for await...of, Symbol.asyncIterator,
/// async generator return/throw, combining async and generator patterns.
use super::helpers::run_js;

// ── async function* basics ────────────────────────────────────────────────────

#[test]
fn async_generator_yields_promises() {
    assert_eq!(
        run_js(
            r#"
async function* gen() {
    yield 1;
    yield 2;
    yield 3;
}
async function main() {
    const results = [];
    for await (const v of gen()) {
        results.push(v);
    }
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn async_generator_can_await_inside() {
    assert_eq!(
        run_js(
            r#"
function delay(v) { return Promise.resolve(v * 10); }
async function* gen() {
    yield await delay(1);
    yield await delay(2);
    yield await delay(3);
}
async function main() {
    const results = [];
    for await (const v of gen()) results.push(v);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["10,20,30"]
    );
}

#[test]
fn async_generator_with_return_value() {
    assert_eq!(
        run_js(
            r#"
async function* gen() {
    yield "a";
    return "final";
}
async function main() {
    const g = gen();
    const r1 = await g.next();
    const r2 = await g.next();
    const r3 = await g.next();
    console.log(r1.value + "," + r1.done);
    console.log(r2.value + "," + r2.done);
    console.log(r3.value + "," + r3.done);
}
main();
"#
        ),
        vec!["a,false", "final,true", "undefined,true"]
    );
}

#[test]
fn async_generator_next_accepts_value() {
    assert_eq!(
        run_js(
            r#"
async function* gen() {
    const x = yield "first";
    yield "second:" + x;
}
async function main() {
    const g = gen();
    await g.next();
    const r = await g.next("hello");
    console.log(r.value);
}
main();
"#
        ),
        vec!["second:hello"]
    );
}

// ── for await...of ────────────────────────────────────────────────────────────

#[test]
fn for_await_of_array_of_promises() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const promises = [Promise.resolve(1), Promise.resolve(2), Promise.resolve(3)];
    const results = [];
    for await (const v of promises) results.push(v);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn for_await_of_with_break() {
    assert_eq!(
        run_js(
            r#"
async function* naturals() {
    let n = 1;
    while (true) yield n++;
}
async function main() {
    const results = [];
    for await (const v of naturals()) {
        if (v > 5) break;
        results.push(v);
    }
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

// ── Symbol.asyncIterator ──────────────────────────────────────────────────────

#[test]
fn custom_async_iterable_via_symbol_async_iterator() {
    assert_eq!(
        run_js(
            r#"
const asyncIterable = {
    [Symbol.asyncIterator]() {
        let i = 0;
        const data = [10, 20, 30];
        return {
            async next() {
                if (i < data.length) return { value: data[i++], done: false };
                return { value: undefined, done: true };
            }
        };
    }
};
async function main() {
    const results = [];
    for await (const v of asyncIterable) results.push(v);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["10,20,30"]
    );
}

// ── async generator error handling ────────────────────────────────────────────

#[test]
fn async_generator_try_catch_handles_throw() {
    assert_eq!(
        run_js(
            r#"
async function* gen() {
    try {
        yield 1;
        yield 2;
    } catch (e) {
        yield "caught:" + e.message;
    }
}
async function main() {
    const g = gen();
    await g.next();
    const r = await g.throw(new Error("boom"));
    console.log(r.value);
}
main();
"#
        ),
        vec!["caught:boom"]
    );
}

#[test]
fn async_generator_finally_runs_on_return() {
    assert_eq!(
        run_js(
            r#"
const log = [];
async function* gen() {
    try {
        yield 1;
    } finally {
        log.push("finally");
    }
}
async function main() {
    const g = gen();
    await g.next();
    await g.return();
    console.log(log.join(","));
}
main();
"#
        ),
        vec!["finally"]
    );
}

// ── async generator pipeline ──────────────────────────────────────────────────

#[test]
fn async_generator_pipeline_map_filter() {
    assert_eq!(
        run_js(
            r#"
async function* range(start, end) {
    for (let i = start; i <= end; i++) yield i;
}
async function* asyncMap(iter, fn) {
    for await (const v of iter) yield fn(v);
}
async function* asyncFilter(iter, pred) {
    for await (const v of iter) if (pred(v)) yield v;
}
async function main() {
    const evensDoubled = asyncFilter(
        asyncMap(range(1, 6), x => x * 2),
        x => x > 4
    );
    const results = [];
    for await (const v of evensDoubled) results.push(v);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["6,8,10,12"]
    );
}

// ── async generator with delegation ──────────────────────────────────────────

#[test]
fn async_generator_yield_star_delegates_to_sync_iterable() {
    assert_eq!(
        run_js(
            r#"
async function* gen() {
    yield* [1, 2, 3];
    yield 4;
}
async function main() {
    const results = [];
    for await (const v of gen()) results.push(v);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["1,2,3,4"]
    );
}

#[test]
fn async_generator_yield_star_delegates_to_async_generator() {
    assert_eq!(
        run_js(
            r#"
async function* inner() {
    yield "x";
    yield "y";
}
async function* outer() {
    yield "start";
    yield* inner();
    yield "end";
}
async function main() {
    const results = [];
    for await (const v of outer()) results.push(v);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["start,x,y,end"]
    );
}

// ── async generator return() ──────────────────────────────────────────────────

#[test]
fn async_generator_return_method_finishes_early() {
    assert_eq!(
        run_js(
            r#"
async function* gen() {
    yield 1;
    yield 2;
    yield 3;
}
async function main() {
    const g = gen();
    const r1 = await g.next();
    const ret = await g.return(99);
    const r2 = await g.next();
    console.log(r1.value);
    console.log(ret.value);
    console.log(r2.done);
}
main();
"#
        ),
        vec!["1", "99", "true"]
    );
}

// ── lazy async generation ─────────────────────────────────────────────────────

#[test]
fn async_generator_lazy_fetch_simulation() {
    assert_eq!(
        run_js(
            r#"
const pages = ["page1", "page2", "page3"];
async function* fetchPages() {
    for (const page of pages) {
        const data = await Promise.resolve(page.toUpperCase());
        yield data;
    }
}
async function main() {
    const results = [];
    for await (const page of fetchPages()) results.push(page);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["PAGE1,PAGE2,PAGE3"]
    );
}

// ── combining sync and async iteration ───────────────────────────────────────

#[test]
fn async_generator_wraps_sync_generator() {
    assert_eq!(
        run_js(
            r#"
function* syncGen() {
    yield 1; yield 2; yield 3;
}
async function* asyncWrap(iter) {
    for (const v of iter) yield await Promise.resolve(v * 10);
}
async function main() {
    const results = [];
    for await (const v of asyncWrap(syncGen())) results.push(v);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["10,20,30"]
    );
}

// ── next() receives values into yield expression ──────────────────────────────

#[test]
fn async_generator_accumulator_pattern() {
    assert_eq!(
        run_js(
            r#"
async function* accumulate() {
    let sum = 0;
    while (true) {
        const val = yield sum;
        if (val === null) return;
        sum += val;
    }
}
async function main() {
    const g = accumulate();
    await g.next();
    await g.next(10);
    await g.next(20);
    const r = await g.next(30);
    console.log(r.value);
}
main();
"#
        ),
        vec!["60"]
    );
}
