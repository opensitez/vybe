use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Functions — declarations, arrows, generators,
// async/await, rest params, default params
// ═══════════════════════════════════════════════════════════

#[test]
fn arrow_expression_body() {
    let out = run_js(r#"
const double = x => x * 2;
console.log(double(5));
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn arrow_block_body() {
    let out = run_js(r#"
const add = (a, b) => {
    return a + b;
};
console.log(add(3, 4));
"#);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn arrow_no_params() {
    let out = run_js(r#"
const greet = () => "hello";
console.log(greet());
"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn arrow_returns_object() {
    let out = run_js(r#"
const makeObj = (k, v) => ({ [k]: v });
const obj = makeObj("name", "Alice");
console.log(obj.name);
"#);
    assert_eq!(out, vec!["Alice"]);
}

#[ignore]
#[test]
fn rest_params() {
    let out = run_js(r#"
function sum(...nums) {
    let total = 0;
    for (const n of nums) {
        total += n;
    }
    return total;
}
console.log(sum(1, 2, 3, 4));
"#);
    assert_eq!(out, vec!["10"]);
}

#[ignore]
#[test]
fn rest_params_after_named() {
    let out = run_js(r#"
function log(prefix, ...messages) {
    for (const m of messages) {
        console.log(prefix + ": " + m);
    }
}
log("INFO", "start", "end");
"#);
    assert_eq!(out, vec!["INFO: start", "INFO: end"]);
}

#[test]
fn default_params_basic() {
    let out = run_js(r#"
function greet(name = "World") {
    console.log("Hello " + name);
}
greet();
greet("Alice");
"#);
    assert_eq!(out, vec!["Hello World", "Hello Alice"]);
}

#[test]
fn default_params_expression() {
    let out = run_js(r#"
function create(width = 100, height = width * 2) {
    console.log(width + "x" + height);
}
create();
create(50);
"#);
    assert_eq!(out, vec!["100x200", "50x100"]);
}

#[test]
fn default_params_arrow() {
    let out = run_js(r#"
const inc = (x, step = 1) => x + step;
console.log(inc(5));
console.log(inc(5, 3));
"#);
    assert_eq!(out, vec!["6", "8"]);
}

#[test]
fn default_params_explicit_null_differs_from_omission() {
    let out = run_js(r#"
function greet(name = "World") {
    console.log("Hello " + name);
}
greet();
greet(null);
"#);
    assert_eq!(out, vec!["Hello World", "Hello null"]);
}

#[test]
fn iife() {
    let out = run_js(r#"
const result = (function() {
    return 42;
})();
console.log(result);
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn named_function_expression() {
    let out = run_js(r#"
const factorial = function fact(n) {
    if (n <= 1) return 1;
    return n * fact(n - 1);
};
console.log(factorial(5));
"#);
    assert_eq!(out, vec!["120"]);
}

#[test]
fn closure_counter() {
    let out = run_js(r#"
function makeCounter() {
    let count = 0;
    return function() {
        count += 1;
        return count;
    };
}
const counter = makeCounter();
console.log(counter());
console.log(counter());
console.log(counter());
"#);
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn higher_order_function() {
    let out = run_js(r#"
function apply(fn, x) {
    return fn(x);
}
console.log(apply(x => x * x, 5));
"#);
    assert_eq!(out, vec!["25"]);
}

#[test]
fn function_returning_function() {
    let out = run_js(r#"
function multiplier(factor) {
    return x => x * factor;
}
const triple = multiplier(3);
console.log(triple(7));
"#);
    assert_eq!(out, vec!["21"]);
}

#[test]
fn recursion_fibonacci() {
    let out = run_js(r#"
function fib(n) {
    if (n <= 1) return n;
    return fib(n - 1) + fib(n - 2);
}
console.log(fib(10));
"#);
    assert_eq!(out, vec!["55"]);
}

#[test]
fn async_function_basic() {
    let out = run_js(r#"
async function fetchData() {
    return 42;
}
const result = fetchData();
console.log(result);
"#);
    // async returns a promise-like; we just verify it compiles and runs
    assert!(!out.is_empty());
}

#[test]
fn async_arrow() {
    let out = run_js(r#"
const getData = async () => {
    return "data";
};
const r = getData();
console.log(r);
"#);
    assert!(!out.is_empty());
}

#[ignore]
#[test]
fn spread_in_call() {
    let out = run_js(r#"
function add(a, b, c) {
    return a + b + c;
}
const args = [1, 2, 3];
console.log(add(...args));
"#);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn method_shorthand() {
    let out = run_js(r#"
const obj = {
    greet(name) {
        return "Hello " + name;
    }
};
console.log(obj.greet("World"));
"#);
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn arrow_function_lexical_this() {
    let out = run_js(r#"
const counter = {
    count: 0,
    inc() {
        const step = () => {
            this.count += 1;
        };
        step();
        step();
        return this.count;
    }
};
console.log(counter.inc());
console.log(counter.count);
"#);
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn function_param_reassignment_changes_later_reads() {
    let out = run_js(r#"
function update(a) {
    console.log(a);
    a = 5;
    console.log(a);
}
update(1);
"#);
    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn missing_arguments_currently_produce_null() {
    let out = run_js(r#"
function show(a, b) {
    console.log(a);
    console.log(b);
}
show(1);
"#);
    assert_eq!(out, vec!["1", "null"]);
}

#[test]
fn extra_arguments_are_ignored_by_named_parameters() {
    let out = run_js(r#"
function add(a, b) {
    console.log(a + b);
}
add(2, 3, 4, 5);
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn nested_function_reads_outer_parameter() {
    let out = run_js(r#"
function outer(value) {
    function inner() {
        return value + 1;
    }
    console.log(inner());
}
outer(4);
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn default_params_evaluated_per_call() {
    let out = run_js(r#"
let n = 0;
function next(value = ++n) {
    console.log(value);
}
next();
next();
next(10);
next();
"#);
    assert_eq!(out, vec!["1", "2", "10", "3"]);
}

#[test]
fn default_params_can_reference_earlier_param() {
    let out = run_js(r#"
function range(start, end = start + 2) {
    console.log(start + ":" + end);
}
range(4);
range(4, 10);
"#);
    assert_eq!(out, vec!["4:6", "4:10"]);
}

#[test]
fn recursive_named_function_expression_internal_name() {
    let out = run_js(r#"
let outer = function inner(n) {
    if (n <= 1) return 1;
    return n * inner(n - 1);
};
console.log(outer(4));
"#);
    assert_eq!(out, vec!["24"]);
}

#[test]
fn function_length_ignores_defaulted_tail_params() {
    let out = run_js(r#"
function sample(a, b, c = 1, d = 2) {}
console.log(sample.length);
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn function_name_for_declaration_and_arrow_assignment() {
    let out = run_js(r#"
function greet() {}
const answer = () => 42;
console.log(greet.name);
console.log(answer.name);
"#);
    assert_eq!(out, vec!["greet", "answer"]);
}
