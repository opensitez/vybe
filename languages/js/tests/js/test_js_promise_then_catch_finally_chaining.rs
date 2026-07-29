use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Promise.prototype.then, catch, and finally Chaining Mechanics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_promise_then_chaining_value_transformation() {
    let src = r#"
Promise.resolve(10)
    .then(x => x * 2)
    .then(x => x + 5)
    .then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["25"]);
}

#[test]
fn test_js_promise_then_returning_nested_promise() {
    let src = r#"
Promise.resolve(5)
    .then(x => new Promise(resolve => resolve(x * 3)))
    .then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["15"]);
}

#[test]
fn test_js_promise_catch_rejection_recovery() {
    let src = r#"
Promise.reject("Initial Error")
    .catch(err => {
        console.log("Caught: " + err);
        return "Recovered";
    })
    .then(res => console.log("Next: " + res));
"#;
    assert_eq!(
        run_js(src),
        vec!["Caught: Initial Error", "Next: Recovered"]
    );
}

#[test]
fn test_js_promise_finally_executes_on_fulfillment() {
    let src = r#"
Promise.resolve("Success")
    .finally(() => console.log("Finally Done"))
    .then(res => console.log("Resolved: " + res));
"#;
    assert_eq!(run_js(src), vec!["Finally Done", "Resolved: Success"]);
}

#[test]
fn test_js_promise_finally_executes_on_rejection() {
    let src = r#"
Promise.reject("Failed")
    .finally(() => console.log("Finally Cleanup"))
    .catch(err => console.log("Handled: " + err));
"#;
    assert_eq!(run_js(src), vec!["Finally Cleanup", "Handled: Failed"]);
}

#[test]
fn test_js_promise_finally_returning_rejected_promise_overrides_value() {
    let src = r#"
Promise.resolve("Initial")
    .finally(() => Promise.reject("Finally Error"))
    .catch(err => console.log("Override Error: " + err));
"#;
    assert_eq!(run_js(src), vec!["Override Error: Finally Error"]);
}

#[test]
fn test_js_promise_then_omitted_handlers_passthrough() {
    let src = r#"
Promise.resolve(100)
    .then(null, err => {}) // Fulfill handler omitted
    .then(val => console.log(val));
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_promise_rejection_omitted_catch_passthrough() {
    let src = r#"
Promise.reject("Uncaught")
    .then(val => val) // Reject handler omitted
    .catch(err => console.log("PassThrough: " + err));
"#;
    assert_eq!(run_js(src), vec!["PassThrough: Uncaught"]);
}

#[test]
fn test_js_promise_executor_throw_triggers_rejection() {
    let src = r#"
new Promise((resolve, reject) => {
    throw new Error("Executor Exception");
})
.catch(err => console.log(err.message));
"#;
    assert_eq!(run_js(src), vec!["Executor Exception"]);
}

#[test]
fn test_js_promise_multiple_then_subscribers_on_same_promise() {
    let src = r#"
const p = Promise.resolve("Shared");
p.then(v => console.log("Sub1: " + v));
p.then(v => console.log("Sub2: " + v));
"#;
    assert_eq!(run_js(src), vec!["Sub1: Shared", "Sub2: Shared"]);
}

#[test]
fn test_js_promise_resolve_twice_first_wins() {
    let src = r#"
new Promise((resolve, reject) => {
    resolve("First");
    resolve("Second");
    reject("Third");
})
.then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["First"]);
}

#[test]
fn test_js_promise_reject_twice_first_wins() {
    let src = r#"
new Promise((resolve, reject) => {
    reject("FirstErr");
    resolve("FirstVal");
})
.catch(err => console.log(err));
"#;
    assert_eq!(run_js(src), vec!["FirstErr"]);
}

#[test]
fn test_js_promise_then_throwing_exception_rejects_chain() {
    let src = r#"
Promise.resolve(1)
    .then(x => { throw new Error("ThrowInThen"); })
    .catch(err => console.log(err.message));
"#;
    assert_eq!(run_js(src), vec!["ThrowInThen"]);
}

#[test]
fn test_js_promise_chaining_thenable_object() {
    let src = r#"
const thenable = {
    then(resolve, reject) {
        resolve("ThenableResult");
    }
};
Promise.resolve().then(() => thenable).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["ThenableResult"]);
}

#[test]
fn test_js_promise_finally_preserves_original_rejection() {
    let src = r#"
Promise.reject("OriginalError")
    .finally(() => "Returned String Ignored")
    .catch(err => console.log(err));
"#;
    assert_eq!(run_js(src), vec!["OriginalError"]);
}

#[test]
fn test_js_promise_then_returns_self_rejection_typeerror() {
    let src = r#"
let p;
p = Promise.resolve().then(() => p);
p.catch(err => console.log(err.name));
"#;
    assert_eq!(run_js(src), vec!["TypeError"]);
}

#[test]
fn test_js_promise_async_resolve_delay_propagation() {
    let src = r#"
new Promise(resolve => {
    resolve(new Promise(res2 => res2(42)));
})
.then(v => console.log(v));
"#;
    assert_eq!(run_js(src), vec!["42"]);
}

#[test]
fn test_js_promise_catch_returns_rejected_promise() {
    let src = r#"
Promise.reject("ErrorA")
    .catch(err => Promise.reject("ErrorB"))
    .catch(err => console.log(err));
"#;
    assert_eq!(run_js(src), vec!["ErrorB"]);
}

#[test]
fn test_js_promise_then_non_function_args_ignored() {
    let src = r#"
Promise.resolve("Original")
    .then("not_a_function")
    .then(12345)
    .then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["Original"]);
}

#[test]
fn test_js_promise_deep_chain_accumulation() {
    let src = r#"
let p = Promise.resolve(0);
for (let i = 1; i <= 5; i++) {
    p = p.then(acc => acc + i);
}
p.then(total => console.log(total));
"#;
    assert_eq!(run_js(src), vec!["15"]);
}

#[test]
fn test_js_promise_catch_throwing_rejects_chain() {
    let src = r#"
Promise.reject("init")
    .catch(() => { throw new Error("catch_throw"); })
    .catch(err => console.log(err.message));
"#;
    assert_eq!(run_js(src), vec!["catch_throw"]);
}

