use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Unhandled Rejection Tracking & Error Propagation Semantics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_promise_handled_rejection_prevents_unhandled_error() {
    let src = r#"
const p = Promise.reject("HandledError");
p.catch(err => console.log("Caught: " + err));
"#;
    assert_eq!(run_js(src), vec!["Caught: HandledError"]);
}

#[test]
fn test_js_promise_late_handler_attachment() {
    let src = r#"
const p = Promise.reject("LateError");
// Attaching handler in subsequent microtask
Promise.resolve().then(() => {
    p.catch(err => console.log("Late Handled: " + err));
});
"#;
    assert_eq!(run_js(src), vec!["Late Handled: LateError"]);
}

#[test]
fn test_js_promise_rejection_in_nested_then_chain() {
    let src = r#"
Promise.resolve(10)
    .then(x => { throw new Error("Step 1 Failed"); })
    .then(x => x * 2) // Skipped!
    .catch(err => console.log(err.message));
"#;
    assert_eq!(run_js(src), vec!["Step 1 Failed"]);
}

#[test]
fn test_js_promise_rejection_rethrown_in_catch() {
    let src = r#"
Promise.reject("Initial")
    .catch(err => {
        console.log("Stage 1: " + err);
        throw new Error("Rethrown");
    })
    .catch(err => console.log("Stage 2: " + err.message));
"#;
    assert_eq!(run_js(src), vec!["Stage 1: Initial", "Stage 2: Rethrown"]);
}

#[test]
fn test_js_promise_suppressed_error_in_async_chain() {
    let src = r#"
Promise.reject("ErrorToSuppress")
    .then(null, err => {
        console.log("Suppressed: " + err);
        return "CleanReturn";
    })
    .then(res => console.log("Result: " + res));
"#;
    assert_eq!(
        run_js(src),
        vec!["Suppressed: ErrorToSuppress", "Result: CleanReturn"]
    );
}

#[test]
fn test_js_promise_multiple_catch_branches() {
    let src = r#"
const p = Promise.reject("RootError");
p.catch(err => console.log("Branch 1: " + err));
p.catch(err => console.log("Branch 2: " + err));
"#;
    assert_eq!(
        run_js(src),
        vec!["Branch 1: RootError", "Branch 2: RootError"]
    );
}

#[test]
fn test_js_promise_finally_does_not_swallow_rejection() {
    let src = r#"
Promise.reject("CriticalFailure")
    .finally(() => console.log("Logging Failure"))
    .catch(err => console.log("Caught After Finally: " + err));
"#;
    assert_eq!(
        run_js(src),
        vec!["Logging Failure", "Caught After Finally: CriticalFailure"]
    );
}

#[test]
fn test_js_promise_rejection_with_non_error_object() {
    let src = r#"
Promise.reject({ status: 500, message: "Internal Error" })
    .catch(err => console.log(err.status + ":" + err.message));
"#;
    assert_eq!(run_js(src), vec!["500:Internal Error"]);
}

#[test]
fn test_js_promise_rejection_with_primitive_null() {
    let src = r#"
Promise.reject(null)
    .catch(err => console.log("Rejected Null: " + (err === null)));
"#;
    assert_eq!(run_js(src), vec!["Rejected Null: true"]);
}

#[test]
fn test_js_promise_rejection_with_primitive_number() {
    let src = r#"
Promise.reject(404)
    .catch(err => console.log("Error Code: " + err));
"#;
    assert_eq!(run_js(src), vec!["Error Code: 404"]);
}

#[test]
fn test_js_promise_catch_rejection_returns_undefined() {
    let src = r#"
Promise.reject("Err")
    .catch(err => {}) // Returns undefined
    .then(val => console.log(val === undefined));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_unhandled_rejection_event_simulation() {
    let src = r#"
const unhandledRejections = [];
function onUnhandledRejection(reason, promise) {
    unhandledRejections.push(reason);
}

const p = Promise.reject("SimulatedUnhandled");
// Simulated global handler check before microtask turn ends
onUnhandledRejection("SimulatedUnhandled", p);
p.catch(() => {}); // Prevent actual unhandled process exit

console.log(unhandledRejections.join(","));
"#;
    assert_eq!(run_js(src), vec!["SimulatedUnhandled"]);
}

#[test]
fn test_js_promise_error_stack_trace_preservation() {
    let src = r#"
const err = new Error("StackTest");
Promise.reject(err).catch(e => console.log(e.stack !== undefined));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_rejection_handling_in_then_second_callback() {
    let src = r#"
Promise.reject("DirectReject")
    .then(
        val => console.log("Success: " + val),
        err => console.log("HandledInThen: " + err)
    );
"#;
    assert_eq!(run_js(src), vec!["HandledInThen: DirectReject"]);
}

#[test]
fn test_js_promise_rejection_in_then_second_callback_does_not_catch_itself() {
    let src = r#"
Promise.reject("Err1")
    .then(
        val => val,
        err => { throw new Error("Err2"); }
    )
    .catch(err => console.log("CaughtInNextCatch: " + err.message));
"#;
    assert_eq!(run_js(src), vec!["CaughtInNextCatch: Err2"]);
}

#[test]
fn test_js_promise_rejection_propagation_across_tick_boundaries() {
    let src = r#"
let p = Promise.reject("BoundaryError");
Promise.resolve()
    .then(() => p)
    .catch(err => console.log("CrossTick: " + err));
"#;
    assert_eq!(run_js(src), vec!["CrossTick: BoundaryError"]);
}

#[test]
fn test_js_promise_aggregate_error_flattening() {
    let src = r#"
const aggErr = new AggregateError([new Error("E1"), new Error("E2")], "BulkFailure");
Promise.reject(aggErr).catch(err => {
    console.log(err.message + "|" + err.errors.map(e => e.message).join(","));
});
"#;
    assert_eq!(run_js(src), vec!["BulkFailure|E1,E2"]);
}

#[test]
fn test_js_promise_rejection_with_boolean_false() {
    let src = r#"
Promise.reject(false).catch(err => console.log("Rejected False: " + err));
"#;
    assert_eq!(run_js(src), vec!["Rejected False: false"]);
}

#[test]
fn test_js_promise_rejection_with_symbol() {
    let src = r#"
const sym = Symbol("errSym");
Promise.reject(sym).catch(err => console.log(err === sym));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_rejection_recovery_with_new_promise() {
    let src = r#"
Promise.reject("Fail")
    .catch(err => Promise.resolve("FallbackValue"))
    .then(val => console.log("Recovered: " + val));
"#;
    assert_eq!(run_js(src), vec!["Recovered: FallbackValue"]);
}

#[test]
fn promise_rejection_with_undefined_reason() {
    let src = r#"
Promise.reject(undefined).catch(reason => console.log(reason === undefined));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

