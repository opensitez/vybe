use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Async try-catch-finally & Rejection Propagation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_async_try_catch_caught_rejection() {
    let src = r#"
async function run() {
    try {
        await Promise.reject("HandledAsyncError");
    } catch (e) {
        console.log("Caught: " + e);
    }
}
run();
"#;
    assert_eq!(run_js(src), vec!["Caught: HandledAsyncError"]);
}

#[test]
fn test_js_async_finally_always_runs_on_success() {
    let src = r#"
async function run() {
    try {
        return await Promise.resolve("Success");
    } finally {
        console.log("Cleanup Executed");
    }
}
run().then(res => console.log("Result: " + res));
"#;
    assert_eq!(run_js(src), vec!["Cleanup Executed", "Result: Success"]);
}

#[test]
fn test_js_async_finally_always_runs_on_error() {
    let src = r#"
async function run() {
    try {
        await Promise.reject("Failure");
    } catch (e) {
        console.log("Error Handled");
    } finally {
        console.log("Cleanup Always");
    }
}
run();
"#;
    assert_eq!(run_js(src), vec!["Error Handled", "Cleanup Always"]);
}

#[test]
fn test_js_async_finally_override_return_value() {
    let src = r#"
async function testOverride() {
    try {
        return "TryVal";
    } finally {
        return "FinallyVal"; // Finally return overrides try return!
    }
}
testOverride().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["FinallyVal"]);
}

#[test]
fn test_js_async_finally_throw_overrides_rejection() {
    let src = r#"
async function testThrowOverride() {
    try {
        await Promise.reject("OriginalError");
    } finally {
        throw new Error("FinallyError");
    }
}
testThrowOverride().catch(e => console.log(e.message));
"#;
    assert_eq!(run_js(src), vec!["FinallyError"]);
}

#[test]
fn test_js_async_nested_try_catch_blocks() {
    let src = r#"
async function nested() {
    try {
        try {
            await Promise.reject("InnerError");
        } catch (e) {
            console.log("Inner: " + e);
            throw new Error("RethrownInner");
        }
    } catch (e) {
        console.log("Outer: " + e.message);
    }
}
nested();
"#;
    assert_eq!(
        run_js(src),
        vec!["Inner: InnerError", "Outer: RethrownInner"]
    );
}

#[test]
fn test_js_async_try_catch_unhandled_rejection_propagates_to_returned_promise() {
    let src = r#"
async function failUncaught() {
    await Promise.reject("UncaughtInAsync");
}
failUncaught().catch(e => console.log("Caller Caught: " + e));
"#;
    assert_eq!(run_js(src), vec!["Caller Caught: UncaughtInAsync"]);
}

#[test]
fn test_js_async_try_catch_finally_order_of_execution() {
    let src = r#"
async function checkOrder() {
    const log = [];
    try {
        log.push("Try");
        await Promise.resolve();
        throw new Error("Err");
    } catch (e) {
        log.push("Catch");
    } finally {
        log.push("Finally");
    }
    return log.join("->");
}
checkOrder().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["Try->Catch->Finally"]);
}

#[test]
fn test_js_async_try_catch_with_custom_error_types() {
    let src = r#"
class ValidationError extends Error {}
class NetworkError extends Error {}

async function process(type) {
    try {
        if (type === "val") throw new ValidationError("Invalid Input");
        if (type === "net") throw new NetworkError("Connection Timeout");
    } catch (e) {
        if (e instanceof ValidationError) console.log("ValError: " + e.message);
        else if (e instanceof NetworkError) console.log("NetError: " + e.message);
    }
}
(async () => {
    await process("val");
    await process("net");
})();
"#;
    assert_eq!(
        run_js(src),
        vec!["ValError: Invalid Input", "NetError: Connection Timeout"]
    );
}

#[test]
fn test_js_async_catch_clause_without_parameter_es2019() {
    let src = r#"
async function optionalCatch() {
    try {
        await Promise.reject("SecretFail");
    } catch {
        console.log("Caught Without Parameter");
    }
}
optionalCatch();
"#;
    assert_eq!(run_js(src), vec!["Caught Without Parameter"]);
}

#[test]
fn test_js_async_catch_destructuring_error_object() {
    let src = r#"
async function destructureError() {
    try {
        const err = new Error("BadFormat");
        err.code = 400;
        throw err;
    } catch ({ message, code }) {
        console.log(`${message}:${code}`);
    }
}
destructureError();
"#;
    assert_eq!(run_js(src), vec!["BadFormat:400"]);
}

#[test]
fn test_js_async_try_catch_finally_with_await_in_finally() {
    let src = r#"
async function awaitInFinally() {
    try {
        return "Data";
    } finally {
        const msg = await Promise.resolve("AsyncCleanupDone");
        console.log(msg);
    }
}
awaitInFinally().then(res => console.log("Returned: " + res));
"#;
    assert_eq!(run_js(src), vec!["AsyncCleanupDone", "Returned: Data"]);
}

#[test]
fn test_js_async_try_catch_finally_await_in_catch() {
    let src = r#"
async function awaitInCatch() {
    try {
        await Promise.reject("InitialFail");
    } catch (e) {
        const recovered = await Promise.resolve("AsyncRecovery");
        return `${e}->${recovered}`;
    }
}
awaitInCatch().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["InitialFail->AsyncRecovery"]);
}

#[test]
fn test_js_async_try_catch_synchronous_throw_before_await() {
    let src = r#"
async function syncThrowFirst() {
    throw new Error("SyncThrow");
    await Promise.resolve();
}
syncThrowFirst().catch(e => console.log(e.message));
"#;
    assert_eq!(run_js(src), vec!["SyncThrow"]);
}

#[test]
fn test_js_async_try_catch_finally_loop_break_cleanup() {
    let src = r#"
async function loopCleanup() {
    for (let i = 1; i <= 3; i++) {
        try {
            if (i === 2) break;
        } finally {
            console.log("Loop Finally " + i);
        }
    }
}
loopCleanup();
"#;
    assert_eq!(run_js(src), vec!["Loop Finally 1", "Loop Finally 2"]);
}

#[test]
fn test_js_async_try_catch_rejection_of_thenable_object() {
    let src = r#"
async function catchThenable() {
    try {
        await {
            then(resolve, reject) { reject("ThenableFail"); }
        };
    } catch (e) {
        console.log("Caught Thenable: " + e);
    }
}
catchThenable();
"#;
    assert_eq!(run_js(src), vec!["Caught Thenable: ThenableFail"]);
}

#[test]
fn test_js_async_try_catch_finally_return_promise_resolution() {
    let src = r#"
async function returnPromiseInTry() {
    try {
        return Promise.resolve("TryPromiseResolved");
    } finally {
        console.log("Finally Completed");
    }
}
returnPromiseInTry().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["Finally Completed", "TryPromiseResolved"]);
}

#[test]
fn test_js_async_try_catch_finally_return_promise_rejection_in_finally() {
    let src = r#"
async function failInFinally() {
    try {
        return "TrySuccess";
    } finally {
        return Promise.reject("FinallyRejectedPromise");
    }
}
failInFinally().catch(e => console.log(e));
"#;
    assert_eq!(run_js(src), vec!["FinallyRejectedPromise"]);
}

#[test]
fn test_js_async_try_catch_aggregate_error_handling() {
    let src = r#"
async function processAll() {
    try {
        await Promise.any([Promise.reject("ErrA"), Promise.reject("ErrB")]);
    } catch (e) {
        console.log(e.name + "|Count=" + e.errors.length);
    }
}
processAll();
"#;
    assert_eq!(run_js(src), vec!["AggregateError|Count=2"]);
}

#[test]
fn test_js_async_try_catch_deep_async_call_stack_unwinding() {
    let src = r#"
async function level3() { throw new Error("DeepError"); }
async function level2() { await level3(); }
async function level1() {
    try {
        await level2();
    } catch (e) {
        console.log("Unwound Stack: " + e.message);
    }
}
level1();
"#;
    assert_eq!(run_js(src), vec!["Unwound Stack: DeepError"]);
}

#[test]
fn test_js_async_finally_rejecting_promise_overrides_try_rejection() {
    let src = r#"
async function fn() {
    try {
        return Promise.reject("try_rej");
    } finally {
        return Promise.reject("fin_rej");
    }
}
fn().catch(e => console.log(e));
"#;
    assert_eq!(run_js(src), vec!["fin_rej"]);
}
