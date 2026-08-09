/// Promise.withResolvers, Promise.any, Promise.race patterns (ES2024)
use super::helpers::run_js;

#[test]
fn promise_any_resolves_first() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const p = await Promise.any([
        Promise.reject("no"),
        Promise.resolve("yes"),
        Promise.resolve("also yes"),
    ]);
    console.log(p);
}
main();
"#
        ),
        vec!["yes"]
    );
}

#[test]
fn promise_any_all_reject_throws_aggregate() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    try {
        await Promise.any([
            Promise.reject("a"),
            Promise.reject("b"),
        ]);
    } catch (e) {
        console.log(e instanceof AggregateError);
        console.log(e.errors.length);
    }
}
main();
"#
        ),
        vec!["true", "2"]
    );
}

#[test]
fn promise_any_empty_rejects_aggregate() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    try {
        await Promise.any([]);
    } catch (e) {
        console.log(e instanceof AggregateError);
    }
}
main();
"#
        ),
        vec!["true"]
    );
}

#[test]
fn promise_with_resolvers_resolve() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const { promise, resolve } = Promise.withResolvers();
    resolve(42);
    const val = await promise;
    console.log(val);
}
main();
"#
        ),
        vec!["42"]
    );
}

#[test]
fn promise_with_resolvers_reject() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const { promise, reject } = Promise.withResolvers();
    reject(new Error("fail"));
    try {
        await promise;
    } catch (e) {
        console.log(e.message);
    }
}
main();
"#
        ),
        vec!["fail"]
    );
}

#[test]
fn promise_with_resolvers_deferred_resolve() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const { promise, resolve } = Promise.withResolvers();
    setTimeout(() => resolve("deferred"), 0);
    const val = await promise;
    console.log(val);
}
main();
"#
        ),
        vec!["deferred"]
    );
}

#[test]
fn promise_race_fastest_wins() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const fast = Promise.resolve("fast");
    const slow = new Promise(resolve => setTimeout(() => resolve("slow"), 100));
    const result = await Promise.race([slow, fast]);
    console.log(result);
}
main();
"#
        ),
        vec!["fast"]
    );
}

#[test]
fn promise_race_rejection_wins_if_first() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const rejected = Promise.reject(new Error("first"));
    const resolved = Promise.resolve("second");
    try {
        await Promise.race([rejected, resolved]);
    } catch (e) {
        console.log(e.message);
    }
}
main();
"#
        ),
        vec!["first"]
    );
}

#[test]
fn promise_all_settled_captures_all() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const results = await Promise.allSettled([
        Promise.resolve(1),
        Promise.reject("err"),
        Promise.resolve(3),
    ]);
    console.log(results[0].status);
    console.log(results[1].status);
    console.log(results[1].reason);
    console.log(results[2].value);
}
main();
"#
        ),
        vec!["fulfilled", "rejected", "err", "3"]
    );
}

#[test]
fn promise_all_rejects_on_first_failure() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    try {
        await Promise.all([
            Promise.resolve(1),
            Promise.reject("boom"),
            Promise.resolve(3),
        ]);
    } catch (e) {
        console.log(e);
    }
}
main();
"#
        ),
        vec!["boom"]
    );
}

#[test]
fn promise_all_empty_resolves_empty_array() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const result = await Promise.all([]);
    console.log(Array.isArray(result));
    console.log(result.length);
}
main();
"#
        ),
        vec!["true", "0"]
    );
}

#[test]
fn promise_resolve_thenable_assimilation() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const thenable = { then(resolve) { resolve(100); } };
    const result = await Promise.resolve(thenable);
    console.log(result);
}
main();
"#
        ),
        vec!["100"]
    );
}

#[test]
fn promise_all_settled_empty_array_resolves_empty() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const res = await Promise.allSettled([]);
    console.log(Array.isArray(res) + "|" + res.length);
}
main();
"#
        ),
        vec!["true|0"]
    );
}
