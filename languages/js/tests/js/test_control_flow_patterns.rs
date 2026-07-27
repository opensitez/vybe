/// Control flow — switch fallthrough, switch with return, for...of patterns
use super::helpers::run_js;

#[test]
fn switch_fallthrough_behavior() {
    assert_eq!(
        run_js(
            r#"
function test(x) {
    let result = "";
    switch (x) {
        case 1:
            result += "1";
            // fallthrough
        case 2:
            result += "2";
            break;
        case 3:
            result += "3";
    }
    return result;
}
console.log(test(1)); // falls through to 2
console.log(test(2));
console.log(test(3));
"#
        ),
        vec!["12", "2", "3"]
    );
}

#[test]
fn switch_default_at_middle() {
    assert_eq!(
        run_js(
            r#"
function test(x) {
    switch (x) {
        case 1: return "one";
        default: return "other";
        case 2: return "two"; // still reachable
    }
}
console.log(test(1));
console.log(test(2));
console.log(test(99));
"#
        ),
        vec!["one", "two", "other"]
    );
}

#[test]
fn switch_uses_strict_equality() {
    assert_eq!(
        run_js(
            r#"
switch ("1") {
    case 1: console.log("number"); break;
    case "1": console.log("string"); break;
    default: console.log("default");
}
"#
        ),
        vec!["string"]
    );
}

#[test]
fn for_of_with_index_via_entries() {
    assert_eq!(
        run_js(
            r#"
const arr = ["a", "b", "c"];
const result = [];
for (const [i, v] of arr.entries()) {
    result.push(i + ":" + v);
}
console.log(result.join(","));
"#
        ),
        vec!["0:a,1:b,2:c"]
    );
}

#[test]
fn for_of_with_destructuring() {
    assert_eq!(
        run_js(
            r#"
const pairs = [["a", 1], ["b", 2], ["c", 3]];
const result = [];
for (const [key, val] of pairs) {
    result.push(key + "=" + val);
}
console.log(result.join(","));
"#
        ),
        vec!["a=1,b=2,c=3"]
    );
}

#[test]
fn for_await_of_in_async() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const promises = [1, 2, 3].map(x => Promise.resolve(x * x));
    const results = [];
    for await (const v of promises) results.push(v);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["1,4,9"]
    );
}

#[test]
fn while_loop_with_complex_condition() {
    assert_eq!(
        run_js(
            r#"
let a = 1, b = 100, count = 0;
while (a < b) {
    a *= 2;
    b -= 10;
    count++;
}
console.log(count);
console.log(a >= b);
"#
        ),
        vec!["6", "true"]
    );
}

#[test]
fn switch_with_no_break_and_return() {
    assert_eq!(
        run_js(
            r#"
function classify(n) {
    switch (true) {
        case n < 0: return "negative";
        case n === 0: return "zero";
        case n < 10: return "small";
        default: return "large";
    }
}
console.log(classify(-5));
console.log(classify(0));
console.log(classify(7));
console.log(classify(100));
"#
        ),
        vec!["negative", "zero", "small", "large"]
    );
}

#[test]
fn for_in_vs_for_of() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2, c: 3 };
const forInKeys = [];
for (const k in obj) forInKeys.push(k);
// for-of doesn't work on plain objects (not iterable)
const arr = [10, 20, 30];
const forOfVals = [];
for (const v of arr) forOfVals.push(v);
console.log(forInKeys.join(","));
console.log(forOfVals.join(","));
"#
        ),
        vec!["a,b,c", "10,20,30"]
    );
}

#[test]
fn conditional_assignment_patterns() {
    assert_eq!(
        run_js(
            r#"
let x;
x = x || 10;      // short-circuit assignment (before logical assign)
console.log(x);
x = x && (x * 2);
console.log(x);
"#
        ),
        vec!["10", "20"]
    );
}

#[test]
fn if_else_with_switch_value() {
    assert_eq!(
        run_js(
            r#"
const n = 1;
let out;
if (n === 0) {
    out = "zero";
} else if (n === 1) {
    out = "one";
} else {
    out = "other";
}
console.log(out);
            "#
        ),
        vec!["one"]
    );
}

#[test]
fn conditional_expression_chain() {
    let src = r#"
const score = 88;
const grade = score >= 90 ? "A" : score >= 80 ? "B" : "C";
console.log(`${grade}:${score > 100}`);
"#;
    assert_eq!(run_js(src), vec!["B:false"]);
}

#[test]
fn while_loop_continue_and_break_together() {
    let src = r#"
let n = 0;
const values = [];
while (n < 6) {
    n++;
    if (n % 2 === 0) continue;
    if (n > 4) break;
    values.push(n);
}
console.log(values.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,3"]);
}

#[test]
fn for_in_uses_object_enumeration_order() {
    let src = r#"
const obj = { a: 1, c: 2, b: 3 };
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,c,b"]);
}

#[test]
fn condition_evaluates_logical_operators_for_truthiness() {
    let src = r#"
const calls = [];
const side_effect = () => {
    calls.push("side");
    return true;
};

if (0 && side_effect()) {
    calls.push("if-true");
} else {
    calls.push("if-false");
}

if (1 || side_effect()) {
    calls.push("or-true");
}

console.log(calls.join(","));
"#;
    assert_eq!(run_js(src), vec!["if-false,or-true"]);
}

#[test]
fn if_condition_uses_assigned_value_for_truthiness() {
    let src = r#"
let x = 0;
const out = [];

if ((x = 7)) {
    out.push("if");
}
if ((x = 0)) {
    out.push("should-not-run");
} else {
    out.push("else");
}

console.log(out.join("|"));
console.log(x);
"#;
    assert_eq!(run_js(src), vec!["if|else", "0"]);
}

#[test]
fn if_condition_uses_comma_expression_evaluation_order() {
    let src = r#"
const trace = [];
const mark = (name, value) => {
    trace.push(name);
    return value;
};

let x = 0;
if (mark("lhs", x = 5), mark("rhs", x > 2)) {
    trace.push("then");
}

console.log(trace.join(","));
console.log(x);
"#;
    assert_eq!(run_js(src), vec!["lhs,rhs,then", "5"]);
}

#[test]
fn while_loop_condition_increments_and_stops() {
    let src = r#"
let n = 0;
let iterations = 0;
while ((n += 1) < 4) {
    iterations++;
}
console.log(iterations);
console.log(n);
"#;
    assert_eq!(run_js(src), vec!["3", "4"]);
}

#[test]
fn for_of_skips_values_with_continue() {
    assert_eq!(
        run_js(
            r#"
const seen = [];
for (const n of [1, 2, 3, 4]) {
    if (n % 2 === 0) continue;
    seen.push(n);
}
console.log(seen.join(","));
"#
        ),
        vec!["1,3"]
    );
}
