/// Functional programming patterns — compose, curry, memoize, pipe, partial application

use super::helpers::run_js;

#[test]
fn compose_right_to_left() {
    assert_eq!(run_js(r#"
const compose = (...fns) => x => fns.reduceRight((v, f) => f(v), x);
const add1 = x => x + 1;
const double = x => x * 2;
const square = x => x * x;
const transform = compose(add1, double, square);
// square(3) = 9, double(9) = 18, add1(18) = 19
console.log(transform(3));
"#), vec!["19"]);
}

#[test]
fn pipe_left_to_right() {
    assert_eq!(run_js(r#"
const pipe = (...fns) => x => fns.reduce((v, f) => f(v), x);
const transform = pipe(
    x => x * 2,
    x => x + 1,
    x => x.toString()
);
console.log(transform(5)); // 5*2=10, +1=11, toString="11"
"#), vec!["11"]);
}

#[test]
fn curry_creates_partial_applications() {
    assert_eq!(run_js(r#"
const curry = fn => {
    const arity = fn.length;
    return function curried(...args) {
        if (args.length >= arity) return fn(...args);
        return (...more) => curried(...args, ...more);
    };
};
const add = curry((a, b, c) => a + b + c);
console.log(add(1)(2)(3));
console.log(add(1, 2)(3));
console.log(add(1)(2, 3));
"#), vec!["6", "6", "6"]);
}

#[test]
fn partial_application_fixes_first_args() {
    assert_eq!(run_js(r#"
function partial(fn, ...preset) {
    return (...rest) => fn(...preset, ...rest);
}
function multiply(a, b, c) { return a * b * c; }
const double = partial(multiply, 2);
const triple = partial(multiply, 3);
console.log(double(3, 4));
console.log(triple(2, 5));
"#), vec!["24", "30"]);
}

#[test]
fn memoize_caches_results() {
    assert_eq!(run_js(r#"
function memoize(fn) {
    const cache = new Map();
    return function(...args) {
        const key = JSON.stringify(args);
        if (cache.has(key)) return cache.get(key);
        const result = fn.apply(this, args);
        cache.set(key, result);
        return result;
    };
}
let calls = 0;
const expensiveFn = memoize(x => { calls++; return x * x; });
console.log(expensiveFn(5));
console.log(expensiveFn(5));
console.log(expensiveFn(6));
console.log(calls); // only 2: 5 and 6
"#), vec!["25", "25", "36", "2"]);
}

#[test]
fn once_fn_called_only_once() {
    assert_eq!(run_js(r#"
function once(fn) {
    let called = false, result;
    return function(...args) {
        if (!called) { called = true; result = fn(...args); }
        return result;
    };
}
let n = 0;
const inc = once(() => ++n);
console.log(inc());
console.log(inc());
console.log(inc());
console.log(n);
"#), vec!["1", "1", "1", "1"]);
}

#[test]
fn flip_reverses_arguments() {
    assert_eq!(run_js(r#"
const flip = fn => (a, b, ...rest) => fn(b, a, ...rest);
const subtract = (a, b) => a - b;
const flipped = flip(subtract);
console.log(subtract(10, 3));
console.log(flipped(10, 3));
"#), vec!["7", "-7"]);
}

#[test]
fn tap_runs_side_effect_returns_value() {
    assert_eq!(run_js(r#"
const tap = (fn) => (x) => { fn(x); return x; };
const log = tap(x => console.log("tap: " + x));
const result = [1, 2, 3].map(log);
console.log(result.join(","));
"#), vec!["tap: 1", "tap: 2", "tap: 3", "1,2,3"]);
}

#[test]
fn point_free_style_pipeline() {
    assert_eq!(run_js(r#"
const words = ["Hello", "World", "Foo", "Bar"];
const process = arr => arr
    .map(s => s.toLowerCase())
    .filter(s => s.length > 3)
    .sort()
    .join(",");
console.log(process(words));
"#), vec!["hello,world"]);
}

#[test]
fn functor_map_over_custom_type() {
    assert_eq!(run_js(r#"
class Maybe {
    constructor(val) { this.val = val; }
    static of(val) { return new Maybe(val); }
    map(fn) {
        return this.val == null ? this : Maybe.of(fn(this.val));
    }
    get() { return this.val; }
}
const result = Maybe.of(5)
    .map(x => x * 2)
    .map(x => x + 1)
    .get();
console.log(result);
const nullResult = Maybe.of(null).map(x => x * 2).get();
console.log(nullResult);
"#), vec!["11", "null"]);
}
