/// Async/await error handling — rejection patterns, async try/catch/finally,
/// unhandled rejection, async forEach/map patterns, timeout patterns,
/// sequential vs parallel execution.
use super::helpers::run_js;

// ── basic async try/catch ─────────────────────────────────────────────────────

#[test]
fn async_try_catch_catches_thrown_value() {
    assert_eq!(
        run_js(
            r#"
async function f() {
    try {
        throw new Error("async error");
    } catch (e) {
        return e.message;
    }
}
f().then(v => console.log(v));
"#
        ),
        vec!["async error"]
    );
}

#[test]
fn async_catch_on_awaited_rejection() {
    assert_eq!(
        run_js(
            r#"
async function f() {
    try {
        await Promise.reject(new TypeError("bad"));
    } catch (e) {
        return e instanceof TypeError ? "caught" : "wrong";
    }
}
f().then(v => console.log(v));
"#
        ),
        vec!["caught"]
    );
}

// ── async finally ─────────────────────────────────────────────────────────────

#[test]
fn async_finally_always_runs() {
    assert_eq!(
        run_js(
            r#"
const log = [];
async function withFinally(fail) {
    try {
        if (fail) throw new Error("err");
        return "ok";
    } finally {
        log.push("finally");
    }
}
Promise.all([
    withFinally(false).catch(() => {}),
    withFinally(true).catch(() => {})
]).then(() => console.log(log.join(",")));
"#
        ),
        vec!["finally,finally"]
    );
}

#[test]
fn async_finally_does_not_change_resolved_value() {
    assert_eq!(
        run_js(
            r#"
async function f() {
    try { return 42; }
    finally { /* no return */ }
}
f().then(v => console.log(v));
"#
        ),
        vec!["42"]
    );
}

#[test]
fn async_finally_return_overrides_result() {
    assert_eq!(
        run_js(
            r#"
async function f() {
    try { return "try"; }
    finally { return "finally"; } // overrides
}
f().then(v => console.log(v));
"#
        ),
        vec!["finally"]
    );
}

// ── sequential vs parallel ────────────────────────────────────────────────────

#[test]
fn sequential_async_execution_preserves_order() {
    assert_eq!(
        run_js(
            r#"
const log = [];
async function task(n) {
    log.push("start:" + n);
    await Promise.resolve();
    log.push("end:" + n);
    return n;
}
async function main() {
    await task(1);
    await task(2);
}
main().then(() => console.log(log.join(",")));
"#
        ),
        vec!["start:1,end:1,start:2,end:2"]
    );
}

#[test]
fn parallel_async_with_promise_all() {
    assert_eq!(
        run_js(
            r#"
async function delay(n) {
    await Promise.resolve();
    return n * n;
}
async function main() {
    const results = await Promise.all([delay(2), delay(3), delay(4)]);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["4,9,16"]
    );
}

// ── async iteration patterns ──────────────────────────────────────────────────

#[test]
fn sequential_async_map() {
    assert_eq!(
        run_js(
            r#"
async function asyncMap(arr, fn) {
    const results = [];
    for (const item of arr) {
        results.push(await fn(item));
    }
    return results;
}
asyncMap([1, 2, 3], async x => x * x)
    .then(r => console.log(r.join(",")));
"#
        ),
        vec!["1,4,9"]
    );
}

// ── error recovery ────────────────────────────────────────────────────────────

#[test]
fn async_retry_pattern() {
    assert_eq!(
        run_js(
            r#"
async function withRetry(fn, maxAttempts) {
    let lastError;
    for (let i = 0; i < maxAttempts; i++) {
        try { return await fn(i); }
        catch (e) { lastError = e; }
    }
    throw lastError;
}

let attempt = 0;
async function flakyOp(i) {
    attempt++;
    if (attempt < 3) throw new Error("fail");
    return "success";
}

withRetry(flakyOp, 5).then(r => {
    console.log(r);
    console.log(attempt);
});
"#
        ),
        vec!["success", "3"]
    );
}

// ── async error rethrow ───────────────────────────────────────────────────────

#[test]
fn async_rethrow_unknown_errors() {
    assert_eq!(
        run_js(
            r#"
class NetworkError extends Error {}

async function fetchSafe() {
    try {
        throw new NetworkError("timeout");
    } catch (e) {
        if (e instanceof NetworkError) return null;
        throw e; // rethrow unknown
    }
}

fetchSafe().then(v => console.log(v));
"#
        ),
        vec!["null"]
    );
}

// ── async generator interaction ───────────────────────────────────────────────

#[test]
fn for_await_collects_values() {
    assert_eq!(
        run_js(
            r#"
async function* produce() {
    yield await Promise.resolve(1);
    yield await Promise.resolve(2);
    yield await Promise.resolve(3);
}
async function collect() {
    const results = [];
    for await (const v of produce()) results.push(v);
    return results;
}
collect().then(r => console.log(r.join(",")));
"#
        ),
        vec!["1,2,3"]
    );
}

// ── async timeout pattern ─────────────────────────────────────────────────────

#[test]
fn async_timeout_with_promise_race() {
    assert_eq!(
        run_js(
            r#"
function timeout(ms, reason) {
    return new Promise((_, reject) =>
        setTimeout(() => reject(new Error(reason)), ms)
    );
}
async function withTimeout(fn, ms) {
    return Promise.race([fn(), timeout(ms, "timeout")]);
}

const fast = () => Promise.resolve("done");
withTimeout(fast, 1000).then(v => console.log(v));
"#
        ),
        vec!["done"]
    );
}

// ── chained async operations ──────────────────────────────────────────────────

#[test]
fn async_pipeline_pattern() {
    assert_eq!(
        run_js(
            r#"
const steps = [
    async (x) => x + 1,
    async (x) => x * 2,
    async (x) => x - 3,
];

async function pipeline(input, fns) {
    let value = input;
    for (const fn of fns) value = await fn(value);
    return value;
}

pipeline(5, steps).then(v => console.log(v)); // (5+1)*2-3 = 9
"#
        ),
        vec!["9"]
    );
}

// ── async class methods ───────────────────────────────────────────────────────

#[test]
fn async_method_in_class() {
    assert_eq!(
        run_js(
            r#"
class DataService {
    async fetch(id) {
        await Promise.resolve();
        return { id, data: "result:" + id };
    }
}
const svc = new DataService();
svc.fetch(42).then(r => console.log(r.data));
"#
        ),
        vec!["result:42"]
    );
}

#[test]
fn async_try_catch_thrown_object() {
    assert_eq!(
        run_js(
            r#"
async function f() {
    try {
        throw { code: 500 };
    } catch (e) {
        return e.code;
    }
}
f().then(v => console.log(v));
"#
        ),
        vec!["500"]
    );
}

