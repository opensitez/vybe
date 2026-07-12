/// Promise patterns — chaining, error recovery, async utilities
use super::helpers::run_js;

#[test]
fn promise_chain_transforms_value() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const result = await Promise.resolve(1)
        .then(x => x + 1)
        .then(x => x * 3)
        .then(x => "value: " + x);
    console.log(result);
}
main();
"#
        ),
        vec!["value: 6"]
    );
}

#[test]
fn promise_catch_recovery() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const result = await Promise.reject("err")
        .catch(e => "recovered from: " + e)
        .then(v => v + "!");
    console.log(result);
}
main();
"#
        ),
        vec!["recovered from: err!"]
    );
}

#[test]
fn async_function_returns_promise() {
    assert_eq!(
        run_js(
            r#"
async function f() { return 42; }
const p = f();
console.log(p instanceof Promise);
p.then(v => console.log(v));
"#
        ),
        vec!["true", "42"]
    );
}

#[test]
fn await_unwraps_nested_promise() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const p = Promise.resolve(Promise.resolve(Promise.resolve(42)));
    console.log(await p);
}
main();
"#
        ),
        vec!["42"]
    );
}

#[test]
fn async_sequential_order() {
    assert_eq!(
        run_js(
            r#"
const log = [];
async function step(n) {
    log.push("start " + n);
    await Promise.resolve();
    log.push("end " + n);
    return n;
}
async function main() {
    await step(1);
    await step(2);
    console.log(log.join(","));
}
main();
"#
        ),
        vec!["start 1,end 1,start 2,end 2"]
    );
}

#[test]
fn promise_finally_passes_value_through() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const result = await Promise.resolve(42)
        .finally(() => { /* cleanup */ })
        .then(v => v * 2);
    console.log(result);
}
main();
"#
        ),
        vec!["84"]
    );
}

#[test]
fn promise_all_collects_values_in_order() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const results = await Promise.all([
        Promise.resolve(1),
        Promise.resolve(2),
        Promise.resolve(3),
    ]);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn async_error_propagates() {
    assert_eq!(
        run_js(
            r#"
async function throws() {
    throw new Error("async error");
}
async function main() {
    try {
        await throws();
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
fn promise_then_catches_sync_throw() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const result = await Promise.resolve()
        .then(() => { throw new Error("in then"); })
        .catch(e => "caught: " + e.message);
    console.log(result);
}
main();
"#
        ),
        vec!["caught: in then"]
    );
}

#[test]
fn async_iteration_map() {
    assert_eq!(
        run_js(
            r#"
async function asyncMap(arr, fn) {
    return Promise.all(arr.map(fn));
}
async function main() {
    const results = await asyncMap([1, 2, 3], async x => x * 2);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["2,4,6"]
    );
}

#[test]
fn promise_resolve_is_identity_for_same_realm_promise() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const p = Promise.resolve(42);
    const p2 = Promise.resolve(p); // same promise
    console.log(p === p2);
}
main();
"#
        ),
        vec!["true"]
    );
}
