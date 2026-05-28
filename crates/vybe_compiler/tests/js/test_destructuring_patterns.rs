/// Destructuring with defaults, rest patterns, nested, computed keys,
/// patterns in catch/for-of parameters, swaps, mixed types.

use super::helpers::run_js;

// ── defaults in destructuring ─────────────────────────────────────────────────

#[test]
fn object_destructure_default_expression_evaluates() {
    assert_eq!(run_js(r#"
let calls = 0;
function def() { calls++; return 42; }
const { x = def(), y = def() } = { x: 1 };
console.log(x);   // 1 — def() not called
console.log(y);   // 42 — def() called
console.log(calls); // 1
"#), vec!["1", "42", "1"]);
}

#[test]
fn array_destructure_default_from_previous_element() {
    assert_eq!(run_js(r#"
const [a = 10, b = a * 2] = [5];
console.log(a);
console.log(b); // b = a * 2 = 10
"#), vec!["5", "10"]);
}

#[test]
fn destructure_default_not_applied_to_null() {
    assert_eq!(run_js(r#"
// Default only applied when value is undefined, not null
const { x = "default" } = { x: null };
console.log(x);
"#), vec!["null"]);
}

// ── computed keys in destructuring ────────────────────────────────────────────

#[test]
fn computed_key_destructuring() {
    assert_eq!(run_js(r#"
const key = "name";
const { [key]: value } = { name: "Alice" };
console.log(value);
"#), vec!["Alice"]);
}

#[test]
fn computed_key_with_rename() {
    assert_eq!(run_js(r#"
const prop = "color";
const { [prop]: c = "black" } = { color: "red" };
console.log(c);
"#), vec!["red"]);
}

// ── rest in nested destructuring ──────────────────────────────────────────────

#[test]
fn nested_destructure_with_rest() {
    assert_eq!(run_js(r#"
const { a, b, ...rest } = { a: 1, b: 2, c: 3, d: 4 };
console.log(a);
console.log(b);
console.log(JSON.stringify(rest));
"#), vec!["1", "2", "{\"c\":3,\"d\":4}"]);
}

#[test]
fn array_destructure_rest_in_middle() {
    assert_eq!(run_js(r#"
// Rest must be last
const [first, ...remaining] = [1, 2, 3, 4, 5];
console.log(first);
console.log(remaining.join(","));
"#), vec!["1", "2,3,4,5"]);
}

// ── destructuring in function params ──────────────────────────────────────────

#[test]
fn function_param_destructure_object_default() {
    assert_eq!(run_js(r#"
function greet({ name = "World", greeting = "Hello" } = {}) {
    return `${greeting}, ${name}!`;
}
console.log(greet({ name: "Alice" }));
console.log(greet({ greeting: "Hi" }));
console.log(greet());
"#), vec!["Hello, Alice!", "Hi, World!", "Hello, World!"]);
}

#[test]
fn function_param_destructure_array() {
    assert_eq!(run_js(r#"
function sum([a, b, c = 0]) {
    return a + b + c;
}
console.log(sum([1, 2, 3]));
console.log(sum([4, 5]));
"#), vec!["6", "9"]);
}

// ── catch binding destructure ─────────────────────────────────────────────────

#[test]
fn catch_can_destructure_error() {
    assert_eq!(run_js(r#"
try {
    throw { code: 404, message: "not found" };
} catch ({ code, message }) {
    console.log(code);
    console.log(message);
}
"#), vec!["404", "not found"]);
}

// ── for-of with destructuring ─────────────────────────────────────────────────

#[test]
fn for_of_destructure_nested_arrays() {
    assert_eq!(run_js(r#"
const matrix = [[1, 2], [3, 4], [5, 6]];
const sums = [];
for (const [a, b] of matrix) sums.push(a + b);
console.log(sums.join(","));
"#), vec!["3,7,11"]);
}

#[test]
fn for_of_destructure_object_entries() {
    assert_eq!(run_js(r#"
const scores = { alice: 95, bob: 87, charlie: 92 };
const results = [];
for (const [name, score] of Object.entries(scores)) {
    results.push(`${name}:${score}`);
}
console.log(results.sort().join(","));
"#), vec!["alice:95,bob:87,charlie:92"]);
}

// ── iterator destructuring ─────────────────────────────────────────────────────

#[test]
fn destructure_generator_output() {
    assert_eq!(run_js(r#"
function* range(n) { for (let i = 0; i < n; i++) yield i; }
const [a, b, c] = range(5);
console.log(a);
console.log(b);
console.log(c);
"#), vec!["0", "1", "2"]);
}

// ── swap via destructuring ────────────────────────────────────────────────────

#[test]
fn swap_variables_without_temp() {
    assert_eq!(run_js(r#"
let x = 1, y = 2;
[x, y] = [y, x];
console.log(x);
console.log(y);
"#), vec!["2", "1"]);
}

// ── string destructuring ──────────────────────────────────────────────────────

#[test]
fn destructure_string_as_iterable() {
    assert_eq!(run_js(r#"
const [a, b, c] = "hello";
console.log(a);
console.log(c);
"#), vec!["h", "l"]);
}

// ── nested array in object ────────────────────────────────────────────────────

#[test]
fn deep_nested_mixed_destructuring() {
    assert_eq!(run_js(r#"
const data = {
    user: {
        name: "Bob",
        scores: [10, 20, 30]
    }
};
const { user: { name, scores: [first, , third] } } = data;
console.log(name);
console.log(first);
console.log(third);
"#), vec!["Bob", "10", "30"]);
}

// ── ignoring values ───────────────────────────────────────────────────────────

#[test]
fn destructure_ignoring_elements() {
    assert_eq!(run_js(r#"
const [,, third,, fifth = 50] = [1, 2, 3, 4];
console.log(third);
console.log(fifth);
"#), vec!["3", "50"]);
}
