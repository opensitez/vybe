use super::helpers::run_js;

// ── Promise.any ───────────────────────────────────────────
#[test]
fn promise_any_first_resolved() {
    assert_eq!(
        run_js(
            r#"
Promise.any([
  Promise.reject("err1"),
  Promise.resolve("ok"),
  Promise.resolve("second")
]).then(v => console.log(v));
"#
        ),
        vec!["ok"]
    );
}

#[test]
fn promise_any_all_rejected() {
    assert_eq!(
        run_js(
            r#"
Promise.any([
  Promise.reject("a"),
  Promise.reject("b")
]).catch(e => console.log(e instanceof AggregateError));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn promise_any_aggregate_error_errors() {
    assert_eq!(
        run_js(
            r#"
Promise.any([Promise.reject(1), Promise.reject(2)]).catch(e => {
  console.log(e.errors.length);
});
"#
        ),
        vec!["2"]
    );
}

// ── Promise.allSettled ────────────────────────────────────
#[test]
fn promise_allsettled_mixed() {
    assert_eq!(
        run_js(
            r#"
Promise.allSettled([
  Promise.resolve("a"),
  Promise.reject("b"),
  Promise.resolve("c")
]).then(results => {
  const statuses = results.map(r => r.status).join(",");
  console.log(statuses);
});
"#
        ),
        vec!["fulfilled,rejected,fulfilled"]
    );
}

#[test]
fn promise_allsettled_values_and_reasons() {
    assert_eq!(
        run_js(
            r#"
Promise.allSettled([
  Promise.resolve(42),
  Promise.reject("oops")
]).then(results => {
  console.log(results[0].value);
  console.log(results[1].reason);
});
"#
        ),
        vec!["42", "oops"]
    );
}

// ── Promise chaining ──────────────────────────────────────
#[test]
fn promise_then_chain_transforms() {
    assert_eq!(
        run_js(
            r#"
Promise.resolve(1)
  .then(x => x + 1)
  .then(x => x * 3)
  .then(x => console.log(x));
"#
        ),
        vec!["6"]
    );
}

#[test]
fn promise_then_returns_new_promise() {
    assert_eq!(
        run_js(
            r#"
const p1 = Promise.resolve("a");
const p2 = p1.then(v => v + "b");
p2.then(v => console.log(v));
"#
        ),
        vec!["ab"]
    );
}

#[test]
fn promise_catch_recovers() {
    assert_eq!(
        run_js(
            r#"
Promise.reject("error")
  .catch(e => "recovered: " + e)
  .then(v => console.log(v));
"#
        ),
        vec!["recovered: error"]
    );
}

#[test]
fn promise_finally_runs_on_resolve() {
    assert_eq!(
        run_js(
            r#"
const steps = [];
Promise.resolve("ok")
  .finally(() => steps.push("finally"))
  .then(v => { steps.push(v); console.log(steps.join(",")); });
"#
        ),
        vec!["finally,ok"]
    );
}

#[test]
fn promise_finally_runs_on_reject() {
    assert_eq!(
        run_js(
            r#"
const steps = [];
Promise.reject("err")
  .finally(() => steps.push("finally"))
  .catch(e => { steps.push(e); console.log(steps.join(",")); });
"#
        ),
        vec!["finally,err"]
    );
}

// ── Promise.all ───────────────────────────────────────────
#[test]
fn promise_all_resolves_array() {
    assert_eq!(
        run_js(
            r#"
Promise.all([
  Promise.resolve(1),
  Promise.resolve(2),
  Promise.resolve(3)
]).then(vals => console.log(vals.join(",")));
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn promise_all_short_circuits_on_rejection() {
    assert_eq!(
        run_js(
            r#"
Promise.all([
  Promise.resolve(1),
  Promise.reject("fail"),
  Promise.resolve(3)
]).catch(e => console.log(e));
"#
        ),
        vec!["fail"]
    );
}

// ── Promise.race ──────────────────────────────────────────
#[test]
fn promise_race_already_resolved_wins() {
    assert_eq!(
        run_js(
            r#"
Promise.race([
  Promise.resolve("first"),
  Promise.resolve("second"),
  Promise.resolve("third")
]).then(v => console.log(v));
"#
        ),
        vec!["first"]
    );
}

// ── async/await ───────────────────────────────────────────
#[test]
fn async_await_basic() {
    assert_eq!(
        run_js(
            r#"
async function fetchValue() { return 42; }
async function main() {
  const v = await fetchValue();
  console.log(v);
}
main();
"#
        ),
        vec!["42"]
    );
}

#[test]
fn async_await_error_handling() {
    assert_eq!(
        run_js(
            r#"
async function failing() { throw new Error("async error"); }
async function main() {
  try {
    await failing();
  } catch (e) {
    console.log(e.message);
  }
}
main();
"#
        ),
        vec!["async error"]
    );
}

#[test]
fn async_await_sequential_execution() {
    assert_eq!(
        run_js(
            r#"
async function step(n) { return n * 2; }
async function main() {
  const a = await step(1);
  const b = await step(a);
  const c = await step(b);
  console.log(c);
}
main();
"#
        ),
        vec!["8"]
    );
}

#[test]
fn async_returns_promise() {
    assert_eq!(
        run_js(
            r#"
async function f() { return "hello"; }
const p = f();
console.log(p instanceof Promise);
p.then(v => console.log(v));
"#
        ),
        vec!["true", "hello"]
    );
}

#[test]
fn async_await_promise_all_parallel() {
    assert_eq!(
        run_js(
            r#"
async function double(x) { return x * 2; }
async function main() {
  const [a, b, c] = await Promise.all([double(1), double(2), double(3)]);
  console.log(a, b, c);
}
main();
"#
        ),
        vec!["2 4 6"]
    );
}

#[test]
fn async_iife() {
    assert_eq!(
        run_js(
            r#"
(async () => {
  const v = await Promise.resolve("iife");
  console.log(v);
})();
"#
        ),
        vec!["iife"]
    );
}

// ── Promise.resolve/reject ────────────────────────────────
#[test]
fn promise_resolve_thenable() {
    assert_eq!(
        run_js(
            r#"
const thenable = { then(resolve) { resolve(42); } };
Promise.resolve(thenable).then(v => console.log(v));
"#
        ),
        vec!["42"]
    );
}

#[test]
fn promise_resolve_already_promise() {
    assert_eq!(
        run_js(
            r#"
const p = Promise.resolve("original");
const p2 = Promise.resolve(p);
console.log(p === p2);
"#
        ),
        vec!["true"]
    );
}

// ── Error propagation ─────────────────────────────────────
#[test]
fn promise_unhandled_caught_by_chain() {
    assert_eq!(
        run_js(
            r#"
Promise.resolve()
  .then(() => { throw new TypeError("bad type"); })
  .catch(e => console.log(e.constructor.name + ": " + e.message));
"#
        ),
        vec!["TypeError: bad type"]
    );
}

#[test]
fn async_await_rethrow_pattern() {
    assert_eq!(
        run_js(
            r#"
async function withRetry(fn) {
  try {
    return await fn();
  } catch (e) {
    return "fallback: " + e.message;
  }
}
async function main() {
  const r = await withRetry(() => { throw new Error("fail"); });
  console.log(r);
}
main();
"#
        ),
        vec!["fallback: fail"]
    );
}

#[test]
fn promise_chaining_async_values() {
    assert_eq!(
        run_js(
            r#"
async function main() {
  const result = await Promise.resolve(10)
    .then(async x => x + await Promise.resolve(5))
    .then(x => x * 2);
  console.log(result);
}
main();
"#
        ),
        vec!["30"]
    );
}

#[test]
fn promise_all_empty_array() {
    assert_eq!(
        run_js(
            r#"
Promise.all([]).then(v => console.log(v.length));
"#
        ),
        vec!["0"]
    );
}

#[test]
fn promise_allsettled_empty_array() {
    assert_eq!(
        run_js(
            r#"
Promise.allSettled([]).then(v => console.log(v.length));
"#
        ),
        vec!["0"]
    );
}
