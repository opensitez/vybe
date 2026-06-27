/// Promise microtask queue, executor, chaining order, unhandled rejection,
/// scheduler patterns, thenable interop, flattening, resolve/reject races.
use super::helpers::run_js;

// ── executor runs synchronously ───────────────────────────────────────────────

#[test]
fn promise_executor_runs_synchronously() {
    assert_eq!(
        run_js(
            r#"
const log = [];
new Promise((resolve) => {
    log.push("executor");
    resolve();
});
log.push("after");
// microtasks haven't run yet, but executor already ran
console.log(log.join(","));
"#
        ),
        vec!["executor,after"]
    );
}

// ── microtask ordering ────────────────────────────────────────────────────────

#[test]
fn then_callbacks_run_after_current_task() {
    assert_eq!(
        run_js(
            r#"
const log = [];
Promise.resolve().then(() => log.push("microtask"));
log.push("sync");
// After current sync code, microtask runs
Promise.resolve().then(() => console.log(log.join(",")));
"#
        ),
        vec!["microtask,sync"]
    );
}

#[test]
fn multiple_then_callbacks_ordered() {
    assert_eq!(
        run_js(
            r#"
const log = [];
const p = Promise.resolve();
p.then(() => log.push("1"));
p.then(() => log.push("2"));
p.then(() => log.push("3"));
p.then(() => console.log(log.join(",")));
"#
        ),
        vec!["1,2,3"]
    );
}

// ── flattening ────────────────────────────────────────────────────────────────

#[test]
fn returning_promise_from_then_flattens() {
    assert_eq!(
        run_js(
            r#"
Promise.resolve(1)
    .then(v => Promise.resolve(v + 1))
    .then(v => Promise.resolve(v + 1))
    .then(v => console.log(v));
"#
        ),
        vec!["3"]
    );
}

#[test]
fn resolve_with_promise_adopts_state() {
    assert_eq!(
        run_js(
            r#"
const inner = new Promise(resolve => setTimeout(() => resolve("inner"), 0));
Promise.resolve(inner).then(v => console.log(v));
"#
        ),
        vec!["inner"]
    );
}

// ── chaining order ────────────────────────────────────────────────────────────

#[test]
fn then_chain_transforms_value_step_by_step() {
    assert_eq!(
        run_js(
            r#"
Promise.resolve(5)
    .then(n => n * 2)
    .then(n => n + 3)
    .then(n => n.toString())
    .then(s => console.log(s));
"#
        ),
        vec!["13"]
    );
}

// ── rejection propagation ─────────────────────────────────────────────────────

#[test]
fn rejection_skips_then_reaches_catch() {
    assert_eq!(
        run_js(
            r#"
Promise.reject(new Error("fail"))
    .then(v => "unreachable")
    .catch(e => console.log("caught:" + e.message));
"#
        ),
        vec!["caught:fail"]
    );
}

#[test]
fn throw_in_then_becomes_rejection() {
    assert_eq!(
        run_js(
            r#"
Promise.resolve()
    .then(() => { throw new Error("thrown"); })
    .catch(e => console.log(e.message));
"#
        ),
        vec!["thrown"]
    );
}

#[test]
fn catch_recovery_continues_chain() {
    assert_eq!(
        run_js(
            r#"
Promise.reject("bad")
    .catch(e => "recovered")
    .then(v => console.log(v));
"#
        ),
        vec!["recovered"]
    );
}

// ── finally in promise chains ─────────────────────────────────────────────────

#[test]
fn promise_finally_does_not_change_value() {
    assert_eq!(
        run_js(
            r#"
Promise.resolve(42)
    .finally(() => { /* side effect only */ })
    .then(v => console.log(v));
"#
        ),
        vec!["42"]
    );
}

#[test]
fn promise_finally_runs_on_rejection_does_not_suppress() {
    assert_eq!(
        run_js(
            r#"
Promise.reject(new Error("x"))
    .finally(() => { /* runs but doesn't swallow */ })
    .catch(e => console.log(e.message));
"#
        ),
        vec!["x"]
    );
}

// ── Promise.resolve edge cases ────────────────────────────────────────────────

#[test]
fn promise_resolve_with_non_thenable_wraps_it() {
    assert_eq!(
        run_js(
            r#"
Promise.resolve(42).then(v => console.log(v));
"#
        ),
        vec!["42"]
    );
}

#[test]
fn promise_resolve_of_promise_is_identity() {
    assert_eq!(
        run_js(
            r#"
const p = Promise.resolve("same");
const p2 = Promise.resolve(p);
console.log(p === p2);
"#
        ),
        vec!["true"]
    );
}

// ── Promise.reject ────────────────────────────────────────────────────────────

#[test]
fn promise_reject_wraps_reason() {
    assert_eq!(
        run_js(
            r#"
Promise.reject(42).catch(r => console.log(r));
"#
        ),
        vec!["42"]
    );
}

// ── async/await interaction with promises ─────────────────────────────────────

#[test]
fn async_await_awaits_promise_chain() {
    assert_eq!(
        run_js(
            r#"
async function compute() {
    const v1 = await Promise.resolve(10);
    const v2 = await Promise.resolve(v1 * 3);
    return v2;
}
compute().then(v => console.log(v));
"#
        ),
        vec!["30"]
    );
}

#[test]
fn async_try_catch_handles_rejection() {
    assert_eq!(
        run_js(
            r#"
async function run() {
    try {
        await Promise.reject(new Error("async fail"));
    } catch (e) {
        console.log("caught:" + e.message);
    }
}
run();
"#
        ),
        vec!["caught:async fail"]
    );
}

// ── thenable interop ──────────────────────────────────────────────────────────

#[test]
fn thenable_object_is_assimilated_by_promise_resolve() {
    assert_eq!(
        run_js(
            r#"
const thenable = {
    then(resolve) { resolve(99); }
};
Promise.resolve(thenable).then(v => console.log(v));
"#
        ),
        vec!["99"]
    );
}

// ── promise combinators edge cases ────────────────────────────────────────────

#[test]
fn promise_all_resolves_in_order() {
    assert_eq!(
        run_js(
            r#"
const p1 = Promise.resolve(1);
const p2 = Promise.resolve(2);
const p3 = Promise.resolve(3);
Promise.all([p1, p2, p3]).then(values => console.log(values.join(",")));
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn promise_race_fastest_wins() {
    assert_eq!(
        run_js(
            r#"
const fast = Promise.resolve("fast");
const slow = new Promise(r => setTimeout(() => r("slow"), 100));
Promise.race([fast, slow]).then(v => console.log(v));
"#
        ),
        vec!["fast"]
    );
}

#[test]
fn promise_allsettled_includes_all_results() {
    assert_eq!(
        run_js(
            r#"
Promise.allSettled([
    Promise.resolve("ok"),
    Promise.reject("err"),
    Promise.resolve("ok2"),
]).then(results => {
    console.log(results[0].status + ":" + results[0].value);
    console.log(results[1].status + ":" + results[1].reason);
    console.log(results[2].status + ":" + results[2].value);
});
"#
        ),
        vec!["fulfilled:ok", "rejected:err", "fulfilled:ok2"]
    );
}
