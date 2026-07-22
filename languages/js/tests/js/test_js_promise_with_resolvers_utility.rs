use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Promise.withResolvers()` Explicit Resolution Helper (ES2024)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_promise_with_resolvers_structure() {
    let src = r#"
const { promise, resolve, reject } = Promise.withResolvers();
console.log((promise instanceof Promise) + "|" + (typeof resolve === "function") + "|" + (typeof reject === "function"));
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_promise_with_resolvers_external_resolution() {
    let src = r#"
const { promise, resolve } = Promise.withResolvers();
resolve("ExternalValue");
(async () => {
    const val = await promise;
    console.log(val);
})();
"#;
    assert_eq!(run_js(src), vec!["ExternalValue"]);
}

#[test]
fn test_js_promise_with_resolvers_external_rejection() {
    let src = r#"
const { promise, reject } = Promise.withResolvers();
reject("ExternalReason");
(async () => {
    try {
        await promise;
    } catch (reason) {
        console.log("Caught: " + reason);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["Caught: ExternalReason"]);
}

#[test]
fn test_js_promise_with_resolvers_once_settled_subsequent_calls_ignored() {
    let src = r#"
const { promise, resolve, reject } = Promise.withResolvers();
resolve("FirstResolve");
resolve("SecondResolve");
reject("RejectionIgnored");

(async () => {
    const val = await promise;
    console.log(val);
})();
"#;
    assert_eq!(run_js(src), vec!["FirstResolve"]);
}

#[test]
fn test_js_promise_with_resolvers_event_listener_pattern() {
    let src = r#"
class QueueWorker {
    #deferred = Promise.withResolvers();
    get ready() { return this.#deferred.promise; }
    start() { this.#deferred.resolve("WorkerStarted"); }
}
const worker = new QueueWorker();
(async () => {
    const p = worker.ready;
    worker.start();
    console.log(await p);
})();
"#;
    assert_eq!(run_js(src), vec!["WorkerStarted"]);
}

#[test]
fn test_js_promise_with_resolvers_subclassed_promise() {
    let src = r#"
class CustomPromise extends Promise {}
const { promise, resolve } = CustomPromise.withResolvers();
console.log(promise instanceof CustomPromise);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_with_resolvers_resolve_with_promise_thenable() {
    let src = r#"
const { promise, resolve } = Promise.withResolvers();
resolve(Promise.resolve("ChainedPromiseVal"));
(async () => {
    console.log(await promise);
})();
"#;
    assert_eq!(run_js(src), vec!["ChainedPromiseVal"]);
}

#[test]
fn test_js_promise_with_resolvers_resolve_with_rejected_promise() {
    let src = r#"
const { promise, resolve } = Promise.withResolvers();
resolve(Promise.reject("ChainedRejection"));
(async () => {
    try {
        await promise;
    } catch (e) {
        console.log("Caught: " + e);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["Caught: ChainedRejection"]);
}

#[test]
fn test_js_promise_with_resolvers_property_descriptors() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Promise, "withResolvers");
console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["true|false|true"]);
}

#[test]
fn test_js_promise_with_resolvers_length_property() {
    let src = r#"
console.log(Promise.withResolvers.length);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_promise_with_resolvers_name_property() {
    let src = r#"
console.log(Promise.withResolvers.name);
"#;
    assert_eq!(run_js(src), vec!["withResolvers"]);
}

#[test]
fn test_js_promise_with_resolvers_this_non_constructor_throws_typeerror() {
    let src = r#"
try {
    Promise.withResolvers.call(() => {});
} catch (e) {
    console.log("withResolvers Non-Constructor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["withResolvers Non-Constructor TypeError"]);
}

#[test]
fn test_js_promise_with_resolvers_timeout_cancellation_pattern() {
    let src = r#"
const { promise, resolve, reject } = Promise.withResolvers();
const timer = setTimeout(() => reject("TimeoutError"), 1000);

resolve("DataArrivedFast");
clearTimeout(timer);

(async () => {
    console.log(await promise);
})();
"#;
    assert_eq!(run_js(src), vec!["DataArrivedFast"]);
}

#[test]
fn test_js_promise_with_resolvers_resolve_with_undefined() {
    let src = r#"
const { promise, resolve } = Promise.withResolvers();
resolve();
(async () => {
    console.log(await promise === undefined);
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_with_resolvers_reject_with_undefined() {
    let src = r#"
const { promise, reject } = Promise.withResolvers();
reject();
(async () => {
    try {
        await promise;
    } catch (e) {
        console.log(e === undefined);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_promise_with_resolvers_multiple_independent_resolvers() {
    let src = r#"
const r1 = Promise.withResolvers();
const r2 = Promise.withResolvers();
r1.resolve(1);
r2.resolve(2);

(async () => {
    const res = await Promise.all([r1.promise, r2.promise]);
    console.log(res.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_promise_with_resolvers_in_class_field_initializer() {
    let src = r#"
class Task {
    deferred = Promise.withResolvers();
    finish(data) { this.deferred.resolve(data); }
}
const task = new Task();
task.finish("DoneData");
(async () => {
    console.log(await task.deferred.promise);
})();
"#;
    assert_eq!(run_js(src), vec!["DoneData"]);
}

#[test]
fn test_js_promise_with_resolvers_then_chaining() {
    let src = r#"
const { promise, resolve } = Promise.withResolvers();
const chained = promise.then(x => x * 10);
resolve(5);

(async () => {
    console.log(await chained);
})();
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_promise_with_resolvers_catch_chaining() {
    let src = r#"
const { promise, reject } = Promise.withResolvers();
const handled = promise.catch(e => `Handled: ${e}`);
reject("Fail");

(async () => {
    console.log(await handled);
})();
"#;
    assert_eq!(run_js(src), vec!["Handled: Fail"]);
}

#[test]
fn test_js_promise_with_resolvers_finally_chaining() {
    let src = r#"
let finallyCalled = false;
const { promise, resolve } = Promise.withResolvers();
const fin = promise.finally(() => { finallyCalled = true; });
resolve("OK");

(async () => {
    await fin;
    console.log("FinallyCalled=" + finallyCalled);
})();
"#;
    assert_eq!(run_js(src), vec!["FinallyCalled=true"]);
}
