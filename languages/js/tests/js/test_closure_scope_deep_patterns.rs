/// Closure and scope patterns — advanced capture, IIFE, module scope
use super::helpers::run_js;

#[test]
fn closure_in_loop_with_let() {
    assert_eq!(
        run_js(
            r#"
const fns = [];
for (let i = 0; i < 5; i++) {
    fns.push(() => i);
}
console.log(fns.map(f => f()).join(","));
"#
        ),
        vec!["0,1,2,3,4"]
    );
}

#[test]
fn closure_in_loop_var_broken() {
    assert_eq!(
        run_js(
            r#"
const fns = [];
for (var i = 0; i < 5; i++) {
    fns.push(() => i);
}
console.log(fns.map(f => f()).join(","));
"#
        ),
        vec!["5,5,5,5,5"]
    );
}

#[test]
fn iife_private_scope() {
    assert_eq!(
        run_js(
            r#"
const counter = (() => {
    let count = 0;
    return {
        increment() { count++; },
        decrement() { count--; },
        value() { return count; }
    };
})();
counter.increment();
counter.increment();
counter.increment();
counter.decrement();
console.log(counter.value());
"#
        ),
        vec!["2"]
    );
}

#[test]
fn partial_application_closure() {
    assert_eq!(
        run_js(
            r#"
function multiply(a) {
    return function(b) {
        return function(c) {
            return a * b * c;
        };
    };
}
const double = multiply(2);
const times6 = double(3);
console.log(times6(4));
console.log(double(5)(6));
"#
        ),
        vec!["24", "60"]
    );
}

#[test]
fn closure_over_mutable_reference() {
    assert_eq!(
        run_js(
            r#"
function makeAccumulator(initial = 0) {
    let total = initial;
    return {
        add(n) { total += n; return this; },
        subtract(n) { total -= n; return this; },
        result() { return total; } };
}
const acc = makeAccumulator(100);
acc.add(50).add(25).subtract(30);
console.log(acc.result());
"#
        ),
        vec!["145"]
    );
}

#[test]
fn temporal_dead_zone_in_block() {
    assert_eq!(
        run_js(
            r#"
let result = "before";
{
    // let x is not initialized yet (TDZ if accessed here)
    result = "in block";
    let x = 10;
    result = "after let: " + x;
}
console.log(result);
"#
        ),
        vec!["after let: 10"]
    );
}

#[test]
fn function_scope_hoisting() {
    assert_eq!(
        run_js(
            r#"
console.log(hoisted());  // fn declarations hoisted
function hoisted() { return "hoisted"; }
var x = 10;
function useX() { return x; }
console.log(useX());
"#
        ),
        vec!["hoisted", "10"]
    );
}

#[test]
fn nested_function_closure_shared_state() {
    assert_eq!(
        run_js(
            r#"
function makeShared() {
    let shared = [];
    function add(v) { shared.push(v); }
    function get() { return [...shared]; }
    function clear() { shared = []; }
    return { add, get, clear };
}
const { add, get, clear } = makeShared();
add(1); add(2); add(3);
console.log(get().join(","));
clear();
console.log(get().length);
"#
        ),
        vec!["1,2,3", "0"]
    );
}

#[test]
fn closure_with_generator() {
    assert_eq!(
        run_js(
            r#"
function* createIdGen(prefix) {
    let id = 0;
    while (true) yield `${prefix}-${++id}`;
}
const userIds = createIdGen("user");
const postIds = createIdGen("post");
console.log(userIds.next().value);
console.log(userIds.next().value);
console.log(postIds.next().value);
console.log(userIds.next().value);
"#
        ),
        vec!["user-1", "user-2", "post-1", "user-3"]
    );
}

#[test]
fn scope_resolution_lexical() {
    assert_eq!(
        run_js(
            r#"
const x = "global";
function outer() {
    const x = "outer";
    function inner() {
        return x;  // lexically captured
    }
    return inner;
}
const fn = outer();
console.log(fn());  // "outer" not "global"
"#
        ),
        vec!["outer"]
    );
}

#[test]
fn closure_memoization_cache() {
    assert_eq!(
        run_js(
            r#"
function memoize(fn) {
    const cache = new Map();
    return (...args) => {
        const key = JSON.stringify(args);
        if (!cache.has(key)) cache.set(key, fn(...args));
        return cache.get(key);
    };
}
let callCount = 0;
const expensive = memoize((a, b) => { callCount++; return a + b; });
console.log(expensive(1, 2));
console.log(expensive(1, 2));
console.log(expensive(3, 4));
console.log(callCount);
"#
        ),
        vec!["3", "3", "7", "2"]
    );
}

#[test]
fn closure_captures_outer_arguments() {
    assert_eq!(
        run_js(
            r#"
function outer() {
    return () => arguments[0] + arguments[1];
}
const f = outer(10, 20);
console.log(f());
"#
        ),
        vec!["30"]
    );
}
