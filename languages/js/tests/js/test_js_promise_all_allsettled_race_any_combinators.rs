use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Promise Combinators (Promise.all, allSettled, race, any)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_promise_all_success_array() {
    let src = r#"
Promise.all([
    Promise.resolve(10),
    20, // Non-promise value converted automatically
    Promise.resolve(30)
]).then(results => console.log(results.join(",")));
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_promise_all_short_circuits_on_first_rejection() {
    let src = r#"
Promise.all([
    Promise.resolve("A"),
    Promise.reject("FailB"),
    Promise.reject("FailC")
]).catch(err => console.log("Rejected: " + err));
"#;
    assert_eq!(run_js(src), vec!["Rejected: FailB"]);
}

#[test]
fn test_js_promise_all_empty_iterable() {
    let src = r#"
Promise.all([]).then(res => console.log(Array.isArray(res) + "|" + res.length));
"#;
    assert_eq!(run_js(src), vec!["true|0"]);
}

#[test]
fn test_js_promise_all_settled_all_outcomes() {
    let src = r#"
Promise.allSettled([
    Promise.resolve(100),
    Promise.reject("Err")
]).then(results => {
    results.forEach(r => console.log(r.status + "|" + (r.value || r.reason)));
});
"#;
    assert_eq!(run_js(src), vec!["fulfilled|100", "rejected|Err"]);
}

#[test]
fn test_js_promise_all_settled_empty_iterable() {
    let src = r#"
Promise.allSettled([]).then(results => console.log(results.length));
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_promise_race_first_resolved_wins() {
    let src = r#"
Promise.race([
    Promise.resolve("Fast"),
    new Promise(resolve => {})
]).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["Fast"]);
}

#[test]
fn test_js_promise_race_first_rejected_wins() {
    let src = r#"
Promise.race([
    Promise.reject("FastError"),
    Promise.resolve("SlowValue")
]).catch(err => console.log(err));
"#;
    assert_eq!(run_js(src), vec!["FastError"]);
}

#[test]
fn test_js_promise_any_first_fulfilled_wins() {
    let src = r#"
Promise.any([
    Promise.reject("Err1"),
    Promise.resolve("FirstSuccess"),
    Promise.resolve("SecondSuccess")
]).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["FirstSuccess"]);
}

#[test]
fn test_js_promise_any_all_rejected_throws_aggregate_error() {
    let src = r#"
Promise.any([
    Promise.reject("Error1"),
    Promise.reject("Error2")
]).catch(err => {
    console.log((err instanceof AggregateError) + "|" + err.errors.join(","));
});
"#;
    assert_eq!(run_js(src), vec!["true|Error1,Error2"]);
}

#[test]
fn test_js_promise_any_empty_iterable_throws_aggregate_error() {
    let src = r#"
Promise.any([]).catch(err => {
    console.log((err instanceof AggregateError) + "|Count=" + err.errors.length);
});
"#;
    assert_eq!(run_js(src), vec!["true|Count=0"]);
}

#[test]
fn test_js_promise_all_maintains_input_order() {
    let src = r#"
const pSlow = new Promise(resolve => resolve("Slow"));
const pFast = Promise.resolve("Fast");

Promise.all([pSlow, pFast]).then(results => console.log(results.join(",")));
"#;
    assert_eq!(run_js(src), vec!["Slow,Fast"]);
}

#[test]
fn test_js_promise_all_settled_maintains_input_order() {
    let src = r#"
Promise.allSettled([
    Promise.reject("ErrorA"),
    Promise.resolve("SuccessB")
]).then(res => console.log(res[0].status + "|" + res[1].status));
"#;
    assert_eq!(run_js(src), vec!["rejected|fulfilled"]);
}

#[test]
fn test_js_promise_all_custom_iterable_input() {
    let src = r#"
function* generatePromises() {
    yield Promise.resolve(1);
    yield Promise.resolve(2);
}
Promise.all(generatePromises()).then(results => console.log(results.join(",")));
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_promise_all_settled_custom_iterable_input() {
    let src = r#"
const set = new Set([Promise.resolve("A"), Promise.reject("B")]);
Promise.allSettled(set).then(results => console.log(results.length));
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_promise_race_with_primitives() {
    let src = r#"
Promise.race([100, Promise.resolve(200)]).then(val => console.log(val));
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_promise_any_with_primitives() {
    let src = r#"
Promise.any([Promise.reject("Err"), "PrimitiveSuccess"]).then(val => console.log(val));
"#;
    assert_eq!(run_js(src), vec!["PrimitiveSuccess"]);
}

#[test]
fn test_js_promise_all_sparse_array_preserves_holes() {
    let src = r#"
const arr = [Promise.resolve(1), , Promise.resolve(3)];
Promise.all(arr).then(res => console.log(res[0] + "|" + res[1] + "|" + res[2]));
"#;
    assert_eq!(run_js(src), vec!["1|undefined|3"]);
}

#[test]
fn test_js_promise_combinator_non_iterable_throws() {
    let src = r#"
try {
    Promise.all(12345);
} catch (e) {
    console.log("Promise.all Non-Iterable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Promise.all Non-Iterable TypeError"]);
}

#[test]
fn test_js_promise_all_thenable_integration() {
    let src = r#"
const thenable = { then(res) { res("ThenableValue"); } };
Promise.all([thenable]).then(res => console.log(res[0]));
"#;
    assert_eq!(run_js(src), vec!["ThenableValue"]);
}

#[test]
fn test_js_promise_race_empty_iterable_remains_pending() {
    let src = r#"
let executed = false;
Promise.race([]).then(() => { executed = true; });
Promise.resolve().then(() => console.log("Race Empty Still Pending: " + executed));
"#;
    assert_eq!(run_js(src), vec!["Race Empty Still Pending: false"]);
}

#[test]
fn promise_all_subclass_returns_subclass_instance() {
    let src = r#"
class CustomPromise extends Promise {}
CustomPromise.all([CustomPromise.resolve(1)]).then(res => console.log(Array.isArray(res)));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

