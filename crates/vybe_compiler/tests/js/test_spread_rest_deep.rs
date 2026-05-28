/// Spread operator in function calls, array literals, object literals, edge cases

use super::helpers::run_js;

#[test]
fn spread_in_function_call() {
    assert_eq!(run_js(r#"
function sum(a, b, c) { return a + b + c; }
const args = [1, 2, 3];
console.log(sum(...args));
"#), vec!["6"]);
}

#[test]
fn spread_mixed_with_normal_args() {
    assert_eq!(run_js(r#"
function f(a, b, c, d) { return [a, b, c, d].join(","); }
console.log(f(1, ...[2, 3], 4));
"#), vec!["1,2,3,4"]);
}

#[test]
fn spread_array_concat_alternative() {
    assert_eq!(run_js(r#"
const a = [1, 2, 3];
const b = [4, 5, 6];
const combined = [...a, ...b];
console.log(combined.join(","));
"#), vec!["1,2,3,4,5,6"]);
}

#[test]
fn spread_clone_array() {
    assert_eq!(run_js(r#"
const original = [1, 2, 3];
const clone = [...original];
clone.push(4);
console.log(original.length);
console.log(clone.length);
"#), vec!["3", "4"]);
}

#[test]
fn spread_converts_string_to_chars() {
    assert_eq!(run_js(r#"
const chars = [..."hello"];
console.log(chars.join("-"));
"#), vec!["h-e-l-l-o"]);
}

#[test]
fn spread_converts_set() {
    assert_eq!(run_js(r#"
const set = new Set([1, 2, 3, 2, 1]);
const arr = [...set];
console.log(arr.join(","));
"#), vec!["1,2,3"]);
}

#[test]
fn spread_converts_map_to_entries() {
    assert_eq!(run_js(r#"
const map = new Map([["a", 1], ["b", 2]]);
const entries = [...map];
console.log(entries.map(([k,v]) => k+"="+v).join(","));
"#), vec!["a=1,b=2"]);
}

#[test]
fn spread_in_math_max() {
    assert_eq!(run_js(r#"
const nums = [3, 1, 4, 1, 5, 9, 2, 6];
console.log(Math.max(...nums));
console.log(Math.min(...nums));
"#), vec!["9", "1"]);
}

#[test]
fn spread_generator_into_array() {
    assert_eq!(run_js(r#"
function* range(n) { for (let i = 0; i < n; i++) yield i; }
const arr = [...range(5)];
console.log(arr.join(","));
"#), vec!["0,1,2,3,4"]);
}

#[test]
fn rest_collects_remaining_function_args() {
    assert_eq!(run_js(r#"
function log(level, ...messages) {
    return level + ": " + messages.join(", ");
}
console.log(log("INFO", "Starting", "connecting", "ready"));
"#), vec!["INFO: Starting, connecting, ready"]);
}

#[test]
fn spread_insert_in_middle() {
    assert_eq!(run_js(r#"
const start = [1, 2];
const end = [5, 6];
const middle = [3, 4];
const arr = [...start, ...middle, ...end];
console.log(arr.join(","));
"#), vec!["1,2,3,4,5,6"]);
}
