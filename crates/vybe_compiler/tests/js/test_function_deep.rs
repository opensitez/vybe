/// Function features deep — arguments object, callee (sloppy), length/name,
/// toString, default parameters with side effects, rest vs arguments,
/// function as constructor, new.target outside class, bind partial.

use super::helpers::run_js;

// ── arguments object ──────────────────────────────────────────────────────────

#[test]
fn arguments_object_is_array_like() {
    assert_eq!(run_js(r#"
function f() {
    return Array.from(arguments).join(",");
}
console.log(f(1, 2, 3));
"#), vec!["1,2,3"]);
}

#[test]
fn arguments_length_reflects_call() {
    assert_eq!(run_js(r#"
function f(a, b, c) {
    return arguments.length;
}
console.log(f(1, 2));      // 2 args passed
console.log(f(1, 2, 3));   // 3 args passed
console.log(f(1, 2, 3, 4)); // 4 args passed
"#), vec!["2", "3", "4"]);
}

#[test]
fn arguments_vs_rest_params() {
    assert_eq!(run_js(r#"
function withArgs() { return arguments[0]; }
const withRest = (...args) => args[0];
console.log(withArgs(42));
console.log(withRest(42));
"#), vec!["42", "42"]);
}

#[test]
fn arrow_function_has_no_arguments_object() {
    assert_eq!(run_js(r#"
function outer() {
    const inner = () => arguments[0]; // captures outer's arguments
    return inner();
}
console.log(outer(99));
"#), vec!["99"]);
}

// ── function.length ───────────────────────────────────────────────────────────

#[test]
fn function_length_excludes_rest_and_defaults() {
    assert_eq!(run_js(r#"
function f1(a, b, c) {}
function f2(a, b = 1, c) {} // default stops counting
function f3(a, ...rest) {}
console.log(f1.length);
console.log(f2.length);
console.log(f3.length);
"#), vec!["3", "1", "1"]);
}

// ── function.name ─────────────────────────────────────────────────────────────

#[test]
fn function_name_inferred_from_variable() {
    assert_eq!(run_js(r#"
const myFunc = function() {};
const arrow = () => {};
console.log(myFunc.name);
console.log(arrow.name);
"#), vec!["myFunc", "arrow"]);
}

#[test]
fn function_name_inferred_from_object_property() {
    assert_eq!(run_js(r#"
const obj = {
    method() {},
    arrowProp: () => {}
};
console.log(obj.method.name);
console.log(obj.arrowProp.name);
"#), vec!["method", "arrowProp"]);
}

#[test]
fn bound_function_name_has_bound_prefix() {
    assert_eq!(run_js(r#"
function hello() {}
const bound = hello.bind(null);
console.log(bound.name);
"#), vec!["bound hello"]);
}

// ── function.toString ─────────────────────────────────────────────────────────

#[test]
fn function_tostring_contains_source() {
    assert_eq!(run_js(r#"
function add(a, b) { return a + b; }
const src = add.toString();
console.log(src.includes("return a + b"));
"#), vec!["true"]);
}

// ── default parameter with side effect ────────────────────────────────────────

#[test]
fn default_param_evaluated_each_call() {
    assert_eq!(run_js(r#"
let count = 0;
function f(x = ++count) { return x; }
console.log(f());    // 1
console.log(f());    // 2
console.log(f(99));  // 99 — no evaluation
console.log(count);  // 2
"#), vec!["1", "2", "99", "2"]);
}

#[test]
fn default_param_can_reference_previous_param() {
    assert_eq!(run_js(r#"
function f(x, y = x * 2, z = x + y) {
    return `${x},${y},${z}`;
}
console.log(f(3));
console.log(f(3, 10));
"#), vec!["3,6,9", "3,10,13"]);
}

// ── function as constructor ────────────────────────────────────────────────────

#[test]
fn function_constructor_behavior() {
    assert_eq!(run_js(r#"
function Point(x, y) {
    this.x = x;
    this.y = y;
    this.dist = function() { return Math.sqrt(x*x + y*y); };
}
const p = new Point(3, 4);
console.log(p.x);
console.log(p.dist());
console.log(p instanceof Point);
"#), vec!["3", "5", "true"]);
}

// ── new.target outside class ──────────────────────────────────────────────────

#[test]
fn new_target_undefined_in_normal_call() {
    assert_eq!(run_js(r#"
function f() { return new.target; }
console.log(f() === undefined);
console.log(new f() === undefined);
"#), vec!["true", "false"]);
}

#[test]
fn new_target_allows_dual_use_constructor() {
    assert_eq!(run_js(r#"
function Greeter(name) {
    if (new.target === undefined) return new Greeter(name);
    this.name = name;
}
const a = new Greeter("Alice");
const b = Greeter("Bob"); // works without new
console.log(a.name);
console.log(b.name);
"#), vec!["Alice", "Bob"]);
}

// ── bind with partial application ─────────────────────────────────────────────

#[test]
fn bind_partially_applies_arguments() {
    assert_eq!(run_js(r#"
function multiply(a, b) { return a * b; }
const double = multiply.bind(null, 2);
const triple = multiply.bind(null, 3);
console.log(double(5));
console.log(triple(5));
"#), vec!["10", "15"]);
}

#[test]
fn bind_preserves_this_context() {
    assert_eq!(run_js(r#"
const obj = {
    prefix: "Hello",
    greet(name) { return this.prefix + " " + name; }
};
const boundGreet = obj.greet.bind(obj);
const greetAlice = boundGreet.bind(null, "Alice"); // can't override bound this
console.log(greetAlice());
"#), vec!["Hello Alice"]);
}

// ── call and apply ────────────────────────────────────────────────────────────

#[test]
fn call_sets_this_and_passes_args() {
    assert_eq!(run_js(r#"
function greet(greeting) { return greeting + " " + this.name; }
const obj = { name: "Bob" };
console.log(greet.call(obj, "Hi"));
"#), vec!["Hi Bob"]);
}

#[test]
fn apply_passes_args_as_array() {
    assert_eq!(run_js(r#"
function sum(a, b, c) { return a + b + c; }
console.log(sum.apply(null, [1, 2, 3]));
"#), vec!["6"]);
}
