/// Function.prototype — call, apply, bind advanced patterns

use super::helpers::run_js;

#[test]
fn call_with_explicit_this() {
    assert_eq!(run_js(r#"
function greet() { return "Hello, " + this.name; }
const obj = { name: "World" };
console.log(greet.call(obj));
"#), vec!["Hello, World"]);
}

#[test]
fn apply_with_array_args() {
    assert_eq!(run_js(r#"
function sum(a, b, c) { return a + b + c; }
console.log(sum.apply(null, [1, 2, 3]));
"#), vec!["6"]);
}

#[test]
fn bind_creates_new_function() {
    assert_eq!(run_js(r#"
function greet(greeting) { return greeting + ", " + this.name; }
const obj = { name: "Alice" };
const boundGreet = greet.bind(obj);
console.log(boundGreet("Hello"));
console.log(boundGreet("Hi"));
"#), vec!["Hello, Alice", "Hi, Alice"]);
}

#[test]
fn bind_partial_application() {
    assert_eq!(run_js(r#"
function multiply(a, b) { return a * b; }
const double = multiply.bind(null, 2);
const triple = multiply.bind(null, 3);
console.log(double(5));
console.log(triple(4));
"#), vec!["10", "12"]);
}

#[test]
fn bind_preserves_this_in_method() {
    assert_eq!(run_js(r#"
class Timer {
    constructor() { this.count = 0; }
    tick() { this.count++; return this.count; }
}
const t = new Timer();
const tick = t.tick.bind(t);
tick();
tick();
const result = tick();
console.log(result);
console.log(t.count);
"#), vec!["3", "3"]);
}

#[test]
fn call_with_null_this_strict_fn() {
    assert_eq!(run_js(r#"
"use strict";
function whatIsThis() { return this; }
console.log(whatIsThis.call(null));
console.log(whatIsThis.call(undefined));
"#), vec!["null", "undefined"]);
}

#[test]
fn apply_for_variadic_max() {
    assert_eq!(run_js(r#"
const nums = [3, 1, 4, 1, 5, 9, 2, 6];
console.log(Math.max.apply(null, nums));
// Equivalent with spread:
console.log(Math.max(...nums));
"#), vec!["9", "9"]);
}

#[test]
fn bound_function_length_reduces() {
    assert_eq!(run_js(r#"
function f(a, b, c) { return a + b + c; }
console.log(f.length);        // 3
const g = f.bind(null, 1);    // partial: 1 arg bound
console.log(g.length);        // 2
const h = f.bind(null, 1, 2); // 2 args bound
console.log(h.length);        // 1
"#), vec!["3", "2", "1"]);
}

#[test]
fn call_to_borrow_method() {
    assert_eq!(run_js(r#"
// Borrow Array.prototype.slice for array-like objects
function args() { return arguments; }
const argObj = args(1, 2, 3);
const arr = Array.prototype.slice.call(argObj);
console.log(Array.isArray(arr));
console.log(arr.join(","));
"#), vec!["true", "1,2,3"]);
}

#[test]
fn bind_then_new_ignores_this() {
    assert_eq!(run_js(r#"
function Point(x, y) { this.x = x; this.y = y; }
const obj = { name: "ignored" };
const BoundPoint = Point.bind(obj, 1); // bind this and first arg
const p = new BoundPoint(2); // new ignores bound this
console.log(p.x);
console.log(p.y);
console.log(p instanceof Point);
"#), vec!["1", "2", "true"]);
}
