use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Promise.resolve, Promise.reject & Deferred Job Queue Execution
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_promise_resolve_primitive_returns_fulfilled() {
    let src = r#"
Promise.resolve(42).then(v => console.log(v));
"#;
    assert_eq!(run_js(src), vec!["42"]);
}

#[test]
fn test_js_promise_resolve_existing_promise_returns_same_instance() {
    let src = r#"
const p = Promise.resolve("Original");
console.log(Promise.resolve(p) === p);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_reject_existing_promise_wraps_promise_as_reason() {
    let src = r#"
const inner = Promise.resolve("Inner");
const p = Promise.reject(inner);
p.catch(reason => console.log(reason === inner));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_resolve_thenable_adopts_thenable_state() {
    let src = r#"
const thenable = {
    then(onFulfill, onReject) {
        onFulfill("AdoptedThenable");
    }
};
Promise.resolve(thenable).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["AdoptedThenable"]);
}

#[test]
fn test_js_promise_resolve_thenable_rejection_adopts_rejection() {
    let src = r#"
const thenable = {
    then(onFulfill, onReject) {
        onReject("ThenableRejection");
    }
};
Promise.resolve(thenable).catch(err => console.log(err));
"#;
    assert_eq!(run_js(src), vec!["ThenableRejection"]);
}

#[test]
fn test_js_promise_resolve_thenable_throws_exception_in_then() {
    let src = r#"
const badThenable = {
    then() {
        throw new Error("BadThenable");
    }
};
Promise.resolve(badThenable).catch(err => console.log(err.message));
"#;
    assert_eq!(run_js(src), vec!["BadThenable"]);
}

#[test]
fn test_js_promise_with_resolvers_utility_es2024() {
    let src = r#"
const { promise, resolve, reject } = Promise.withResolvers();
promise.then(v => console.log("WithResolvers: " + v));
resolve("SuccessVal");
"#;
    assert_eq!(run_js(src), vec!["WithResolvers: SuccessVal"]);
}

#[test]
fn test_js_promise_with_resolvers_reject_utility() {
    let src = r#"
const { promise, resolve, reject } = Promise.withResolvers();
promise.catch(err => console.log("WithResolversError: " + err));
reject("FailReason");
"#;
    assert_eq!(run_js(src), vec!["WithResolversError: FailReason"]);
}

#[test]
fn test_js_promise_deferred_executor_synchronous_execution() {
    let src = r#"
console.log("Before Promise");
new Promise(resolve => {
    console.log("Inside Executor");
    resolve("Done");
});
console.log("After Promise");
"#;
    assert_eq!(
        run_js(src),
        vec!["Before Promise", "Inside Executor", "After Promise"]
    );
}

#[test]
fn test_js_promise_then_callbacks_always_asynchronous_microtask() {
    let src = r#"
console.log("Start");
Promise.resolve().then(() => console.log("Microtask 1"));
console.log("End");
"#;
    assert_eq!(run_js(src), vec!["Start", "End", "Microtask 1"]);
}

#[test]
fn test_js_promise_resolve_undefined_implicit() {
    let src = r#"
Promise.resolve().then(val => console.log(val === undefined));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_reject_undefined_implicit() {
    let src = r#"
Promise.reject().catch(reason => console.log(reason === undefined));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_subclass_species_constructor() {
    let src = r#"
class MyPromise extends Promise {}
const p = MyPromise.resolve(100);
console.log(p instanceof MyPromise);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_resolve_getter_thenable_property() {
    let src = r#"
let getCount = 0;
const obj = {
    get then() {
        getCount++;
        return (resolve) => resolve("GetterThenable");
    }
};
Promise.resolve(obj).then(res => console.log(res + "|GetCount=" + getCount));
"#;
    assert_eq!(run_js(src), vec!["GetterThenable|GetCount=1"]);
}

#[test]
fn test_js_promise_thenable_called_only_once() {
    let src = r#"
const multicallThenable = {
    then(resolve) {
        resolve("FirstCall");
        resolve("SecondCall");
    }
};
Promise.resolve(multicallThenable).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["FirstCall"]);
}

#[test]
fn test_js_promise_resolve_self_cycle_rejection() {
    let src = r#"
const thenableCycle = {};
thenableCycle.then = (resolve) => resolve(thenableCycle);

Promise.resolve(thenableCycle).catch(err => console.log(err.name));
"#;
    assert_eq!(run_js(src), vec!["TypeError"]);
}

#[test]
fn test_js_promise_executor_ignore_subsequent_reject_after_resolve() {
    let src = r#"
new Promise((resolve, reject) => {
    resolve("Win");
    throw new Error("Ignored Exception After Resolve");
})
.then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["Win"]);
}

#[test]
fn test_js_promise_resolve_symbol_primitive() {
    let src = r#"
const sym = Symbol("test");
Promise.resolve(sym).then(res => console.log(res === sym));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_resolve_bigint_primitive() {
    let src = r#"
Promise.resolve(9007199254740991n).then(res => console.log(res.toString()));
"#;
    assert_eq!(run_js(src), vec!["9007199254740991"]);
}

#[test]
fn test_js_promise_reject_custom_error_subclass() {
    let src = r#"
class CustomAppError extends Error {
    constructor(msg) {
        super(msg);
        this.name = "CustomAppError";
    }
}
Promise.reject(new CustomAppError("App Crash"))
    .catch(err => console.log(err.name + ":" + err.message));
"#;
    assert_eq!(run_js(src), vec!["CustomAppError:App Crash"]);
}
