/// Function edge cases — arguments object, call/apply/bind, named function
/// expressions, function.length/name, IIFE, default param expressions,
/// generator return()/throw(), closure-over-loop with let vs var.
use super::helpers::run_js;

// ── arguments object ──────────────────────────────────────────────────────────

#[test]
fn arguments_object_length_matches_call_args() {
    assert_eq!(
        run_js(
            r#"
function f() { return arguments.length; }
console.log(f(1, 2, 3));
console.log(f());
"#
        ),
        vec!["3", "0"]
    );
}

#[test]
fn arguments_object_indexed_access() {
    assert_eq!(
        run_js(
            r#"
function sum() {
    let total = 0;
    for (let i = 0; i < arguments.length; i++) total += arguments[i];
    return total;
}
console.log(sum(1, 2, 3, 4));
"#
        ),
        vec!["10"]
    );
}

#[test]
fn arguments_spread_to_array() {
    assert_eq!(
        run_js(
            r#"
function toArr() { return Array.from(arguments); }
const a = toArr(10, 20, 30);
console.log(a.join(","));
"#
        ),
        vec!["10,20,30"]
    );
}

#[test]
fn arrow_function_has_no_arguments_object() {
    assert_eq!(
        run_js(
            r#"
function outer() {
    const inner = () => arguments[0];
    return inner();
}
console.log(outer(42));
"#
        ),
        vec!["42"]
    );
}

// ── Function.prototype.call ───────────────────────────────────────────────────

#[test]
fn call_sets_this_context() {
    assert_eq!(
        run_js(
            r#"
function greet() { return "Hello " + this.name; }
const obj = { name: "World" };
console.log(greet.call(obj));
"#
        ),
        vec!["Hello World"]
    );
}

#[test]
fn call_passes_arguments_individually() {
    assert_eq!(
        run_js(
            r#"
function add(a, b, c) { return a + b + c; }
console.log(add.call(null, 1, 2, 3));
"#
        ),
        vec!["6"]
    );
}

#[test]
fn call_with_null_this_in_strict_mode() {
    assert_eq!(
        run_js(
            r#"
function whoami() {
    "use strict";
    return this === null ? "null-this" : "other";
}
console.log(whoami.call(null));
"#
        ),
        vec!["null-this"]
    );
}

// ── Function.prototype.apply ──────────────────────────────────────────────────

#[test]
fn apply_spreads_array_as_args() {
    assert_eq!(
        run_js(
            r#"
function sum(a, b, c) { return a + b + c; }
console.log(sum.apply(null, [10, 20, 30]));
"#
        ),
        vec!["60"]
    );
}

#[test]
fn apply_used_for_math_max_of_array() {
    assert_eq!(
        run_js(
            r#"
const nums = [3, 1, 4, 1, 5, 9, 2, 6];
console.log(Math.max.apply(null, nums));
"#
        ),
        vec!["9"]
    );
}

// ── Function.prototype.bind ───────────────────────────────────────────────────

#[test]
fn bind_fixes_this_permanently() {
    assert_eq!(
        run_js(
            r#"
const obj = { val: 100 };
function getVal() { return this.val; }
const bound = getVal.bind(obj);
console.log(bound());
"#
        ),
        vec!["100"]
    );
}

#[test]
fn bind_partial_application_prepends_args() {
    assert_eq!(
        run_js(
            r#"
function multiply(a, b) { return a * b; }
const double = multiply.bind(null, 2);
console.log(double(5));
console.log(double(10));
"#
        ),
        vec!["10", "20"]
    );
}

#[test]
fn bind_returns_new_function_not_same_ref() {
    assert_eq!(
        run_js(
            r#"
function f() {}
const g = f.bind(null);
console.log(f === g);
"#
        ),
        vec!["false"]
    );
}

#[test]
fn bound_function_ignores_new_this_binding() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 99 };
function getX() { return this.x; }
const bound = getX.bind(obj);
const borrowed = { x: 1 };
console.log(bound.call(borrowed));
"#
        ),
        vec!["99"]
    );
}

// ── function.length and function.name ─────────────────────────────────────────

#[test]
fn function_length_counts_formal_params() {
    assert_eq!(
        run_js(
            r#"
function f(a, b, c) {}
console.log(f.length);
"#
        ),
        vec!["3"]
    );
}

#[test]
fn function_length_excludes_rest_params() {
    assert_eq!(
        run_js(
            r#"
function f(a, b, ...rest) {}
console.log(f.length);
"#
        ),
        vec!["2"]
    );
}

#[test]
fn function_length_excludes_params_after_default() {
    assert_eq!(
        run_js(
            r#"
function f(a, b = 0, c) {}
console.log(f.length);
"#
        ),
        vec!["1"]
    );
}

#[test]
fn function_name_from_declaration() {
    assert_eq!(
        run_js(
            r#"
function myFunc() {}
console.log(myFunc.name);
"#
        ),
        vec!["myFunc"]
    );
}

#[test]
fn function_name_from_variable_assignment() {
    assert_eq!(
        run_js(
            r#"
const myArrow = () => {};
console.log(myArrow.name);
"#
        ),
        vec!["myArrow"]
    );
}

#[test]
fn function_name_named_expression_uses_inner_name() {
    assert_eq!(
        run_js(
            r#"
const f = function namedFn() {};
console.log(f.name);
"#
        ),
        vec!["namedFn"]
    );
}

// ── named function expression scoping ────────────────────────────────────────

#[test]
fn named_function_expression_name_only_visible_inside() {
    assert_eq!(
        run_js(
            r#"
const fib = function fibonacci(n) {
    return n <= 1 ? n : fibonacci(n - 1) + fibonacci(n - 2);
};
console.log(fib(7));
console.log(typeof fibonacci);
"#
        ),
        vec!["13", "undefined"]
    );
}

// ── IIFE patterns ─────────────────────────────────────────────────────────────

#[test]
fn iife_creates_isolated_scope() {
    assert_eq!(
        run_js(
            r#"
const result = (function() {
    const secret = 42;
    return secret * 2;
})();
console.log(result);
console.log(typeof secret);
"#
        ),
        vec!["84", "undefined"]
    );
}

#[test]
fn iife_with_arguments() {
    assert_eq!(
        run_js(
            r#"
const result = (function(a, b) { return a + b; })(10, 20);
console.log(result);
"#
        ),
        vec!["30"]
    );
}

#[test]
fn iife_arrow_function() {
    assert_eq!(
        run_js(
            r#"
const result = (() => 100)();
console.log(result);
"#
        ),
        vec!["100"]
    );
}

// ── default parameter expressions ─────────────────────────────────────────────

#[test]
fn default_param_computed_at_call_time() {
    assert_eq!(
        run_js(
            r#"
let counter = 0;
function f(x = ++counter) { return x; }
console.log(f());
console.log(f());
console.log(f(99));
"#
        ),
        vec!["1", "2", "99"]
    );
}

#[test]
fn default_param_can_reference_earlier_param() {
    assert_eq!(
        run_js(
            r#"
function rect(w, h = w) { return w * h; }
console.log(rect(5));
console.log(rect(3, 4));
"#
        ),
        vec!["25", "12"]
    );
}

#[test]
fn default_param_can_call_function() {
    assert_eq!(
        run_js(
            r#"
function getDefault() { return 42; }
function f(x = getDefault()) { return x; }
console.log(f());
console.log(f(0));
"#
        ),
        vec!["42", "0"]
    );
}

// ── closure over loop variable ────────────────────────────────────────────────

#[test]
fn let_in_loop_creates_fresh_binding_per_iteration() {
    assert_eq!(
        run_js(
            r#"
const fns = [];
for (let i = 0; i < 3; i++) {
    fns.push(() => i);
}
console.log(fns[0]());
console.log(fns[1]());
console.log(fns[2]());
"#
        ),
        vec!["0", "1", "2"]
    );
}

#[test]
fn var_in_loop_shares_binding_across_closures() {
    assert_eq!(
        run_js(
            r#"
const fns = [];
for (var i = 0; i < 3; i++) {
    fns.push(() => i);
}
console.log(fns[0]());
console.log(fns[1]());
console.log(fns[2]());
"#
        ),
        vec!["3", "3", "3"]
    );
}

// ── tail call position ────────────────────────────────────────────────────────

#[test]
fn recursive_accumulator_does_not_overflow_reasonable_depth() {
    assert_eq!(
        run_js(
            r#"
function sum(n, acc = 0) {
    if (n <= 0) return acc;
    return sum(n - 1, acc + n);
}
console.log(sum(100));
"#
        ),
        vec!["5050"]
    );
}

// ── function hoisting ─────────────────────────────────────────────────────────

#[test]
fn function_declaration_hoisted_above_call() {
    assert_eq!(
        run_js(
            r#"
console.log(hoisted());
function hoisted() { return "I was hoisted"; }
"#
        ),
        vec!["I was hoisted"]
    );
}

#[test]
fn function_expression_not_hoisted() {
    assert_eq!(
        run_js(
            r#"
let result;
try {
    result = notHoisted();
} catch (e) {
    result = "error:" + e.constructor.name;
}
var notHoisted = function() { return "late"; };
console.log(result);
"#
        ),
        vec!["error:TypeError"]
    );
}

// ── method shorthand ─────────────────────────────────────────────────────────

#[test]
fn method_shorthand_in_object_literal() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    double(x) { return x * 2; },
    triple(x) { return x * 3; }
};
console.log(obj.double(5));
console.log(obj.triple(5));
"#
        ),
        vec!["10", "15"]
    );
}

// ── getter and setter in object literal ──────────────────────────────────────

#[test]
fn getter_in_object_literal() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    _x: 0,
    get x() { return this._x; },
    set x(v) { this._x = v * 2; }
};
obj.x = 5;
console.log(obj.x);
"#
        ),
        vec!["10"]
    );
}

#[test]
fn getter_computed_each_access() {
    assert_eq!(
        run_js(
            r#"
let count = 0;
const obj = { get id() { return ++count; } };
obj.id; obj.id; obj.id;
console.log(count);
"#
        ),
        vec!["3"]
    );
}

// ── computed property names ───────────────────────────────────────────────────

#[test]
fn computed_property_name_from_expression() {
    assert_eq!(
        run_js(
            r#"
const prefix = "get";
const obj = {
    [prefix + "Name"]() { return "Alice"; },
    [prefix + "Age"]() { return 30; }
};
console.log(obj.getName());
console.log(obj.getAge());
"#
        ),
        vec!["Alice", "30"]
    );
}

#[test]
fn computed_property_name_from_symbol() {
    assert_eq!(
        run_js(
            r#"
const key = Symbol("myKey");
const obj = { [key]: "secret" };
console.log(obj[key]);
"#
        ),
        vec!["secret"]
    );
}
