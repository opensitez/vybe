/// JavaScript Promises: resolve, reject, then, catch, finally,
/// Promise.all, Promise.race, Promise.allSettled, chaining, async/await with promises.
use super::helpers::run_js;

// ===================================================================
// PROMISE BASICS
// ===================================================================

#[test]
fn promise_resolve_then() {
    assert_eq!(
        run_js(
            r#"
Promise.resolve(42).then(v => console.log(v));
"#
        ),
        &["42"]
    );
}

#[test]
fn promise_reject_catch() {
    assert_eq!(
        run_js(
            r#"
Promise.reject("error").catch(e => console.log("caught: " + e));
"#
        ),
        &["caught: error"]
    );
}

#[test]
fn promise_then_chain() {
    assert_eq!(
        run_js(
            r#"
Promise.resolve(1)
  .then(v => v + 1)
  .then(v => v * 3)
  .then(v => console.log(v));
"#
        ),
        &["6"]
    );
}

#[test]
fn promise_catch_and_continue() {
    assert_eq!(
        run_js(
            r#"
Promise.reject("fail")
  .catch(e => "recovered")
  .then(v => console.log(v));
"#
        ),
        &["recovered"]
    );
}

#[test]
fn promise_finally_runs() {
    assert_eq!(
        run_js(
            r#"
Promise.resolve("ok")
  .then(v => console.log(v))
  .finally(() => console.log("done"));
"#
        ),
        &["ok", "done"]
    );
}

#[test]
fn promise_finally_on_reject() {
    assert_eq!(
        run_js(
            r#"
Promise.reject("err")
  .catch(e => console.log("caught"))
  .finally(() => console.log("finally"));
"#
        ),
        &["caught", "finally"]
    );
}

#[test]
fn promise_constructor_resolve() {
    assert_eq!(
        run_js(
            r#"
let p = new Promise((resolve, reject) => {
    resolve("hello");
});
p.then(v => console.log(v));
"#
        ),
        &["hello"]
    );
}

#[test]
fn promise_constructor_reject() {
    assert_eq!(
        run_js(
            r#"
let p = new Promise((resolve, reject) => {
    reject("bad");
});
p.catch(e => console.log("got: " + e));
"#
        ),
        &["got: bad"]
    );
}

// ===================================================================
// PROMISE COMBINATORS
// ===================================================================

#[test]
fn promise_all_basic() {
    assert_eq!(
        run_js(
            r#"
Promise.all([
    Promise.resolve(1),
    Promise.resolve(2),
    Promise.resolve(3)
]).then(values => console.log(values.join(",")));
"#
        ),
        &["1,2,3"]
    );
}

#[test]
fn promise_all_rejects_on_first_failure() {
    assert_eq!(
        run_js(
            r#"
Promise.all([
    Promise.resolve(1),
    Promise.reject("fail"),
    Promise.resolve(3)
]).catch(e => console.log("error: " + e));
"#
        ),
        &["error: fail"]
    );
}

#[test]
fn promise_race_first_wins() {
    assert_eq!(
        run_js(
            r#"
Promise.race([
    Promise.resolve("fast"),
    Promise.resolve("slow")
]).then(v => console.log(v));
"#
        ),
        &["fast"]
    );
}

#[test]
fn promise_allsettled() {
    assert_eq!(
        run_js(
            r#"
Promise.allSettled([
    Promise.resolve("ok"),
    Promise.reject("fail"),
    Promise.resolve("also ok")
]).then(results => {
    results.forEach(r => console.log(r.status));
});
"#
        ),
        &["fulfilled", "rejected", "fulfilled"]
    );
}

// ===================================================================
// ASYNC/AWAIT WITH PROMISES
// ===================================================================

#[test]
fn async_await_resolve() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    let val = await Promise.resolve(42);
    console.log(val);
}
main();
"#
        ),
        &["42"]
    );
}

#[test]
fn async_await_try_catch() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    try {
        let val = await Promise.reject("oops");
    } catch (e) {
        console.log("caught: " + e);
    }
}
main();
"#
        ),
        &["caught: oops"]
    );
}

#[test]
fn async_await_sequential() {
    assert_eq!(
        run_js(
            r#"
async function step(n) {
    return n * 2;
}
async function main() {
    let a = await step(1);
    let b = await step(a);
    let c = await step(b);
    console.log(c);
}
main();
"#
        ),
        &["8"]
    );
}

#[test]
fn async_await_with_promise_all() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    let [a, b, c] = await Promise.all([
        Promise.resolve(10),
        Promise.resolve(20),
        Promise.resolve(30)
    ]);
    console.log(a + b + c);
}
main();
"#
        ),
        &["60"]
    );
}

#[test]
fn promise_constructor_throw_rejects_promise() {
    assert_eq!(
        run_js(
            r#"
new Promise(() => {
    throw new Error("executor_throw");
}).catch(e => console.log(e.message));
"#
        ),
        &["executor_throw"]
    );
}

