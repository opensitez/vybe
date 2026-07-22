use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Closures & Lexical Environment Variable Capture
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_closure_captures_outer_variable_reference() {
    let src = r#"
function createCounter() {
    let count = 0;
    return () => ++count;
}
const counter = createCounter();
console.log(`${counter()}:${counter()}:${counter()}`);
"#;
    assert_eq!(run_js(src), vec!["1:2:3"]);
}

#[test]
fn test_js_closure_multiple_instances_independent_state() {
    let src = r#"
function makeAdder(x) {
    return (y) => x + y;
}
const add5 = makeAdder(5);
const add10 = makeAdder(10);
console.log(add5(3) + "|" + add10(3));
"#;
    assert_eq!(run_js(src), vec!["8|13"]);
}

#[test]
fn test_js_closure_loop_var_shared_binding_issue() {
    let src = r#"
const funcs = [];
for (var i = 0; i < 3; i++) {
    funcs.push(() => i);
}
console.log(funcs.map(f => f()).join(",")); // var shares single binding -> returns 3,3,3
"#;
    assert_eq!(run_js(src), vec!["3,3,3"]);
}

#[test]
fn test_js_closure_loop_let_per_iteration_binding() {
    let src = r#"
const funcs = [];
for (let i = 0; i < 3; i++) {
    funcs.push(() => i);
}
console.log(funcs.map(f => f()).join(",")); // let creates fresh binding per iteration -> returns 0,1,2
"#;
    assert_eq!(run_js(src), vec!["0,1,2"]);
}

#[test]
fn test_js_closure_shared_lexical_environment_mutation() {
    let src = r#"
function createStore() {
    let state = "initial";
    return {
        get: () => state,
        set: (v) => { state = v; }
    };
}
const store = createStore();
console.log(store.get());
store.set("updated");
console.log(store.get());
"#;
    assert_eq!(run_js(src), vec!["initial", "updated"]);
}

#[test]
fn test_js_closure_in_immediately_invoked_function_expression() {
    let src = r#"
const res = (function(secret) {
    return {
        getSecret: () => secret
    };
})("TopSecret");
console.log(res.getSecret());
"#;
    assert_eq!(run_js(src), vec!["TopSecret"]);
}

#[test]
fn test_js_closure_nested_three_levels() {
    let src = r#"
function level1(a) {
    return function level2(b) {
        return function level3(c) {
            return a + b + c;
        };
    };
}
console.log(level1(10)(20)(30));
"#;
    assert_eq!(run_js(src), vec!["60"]);
}

#[test]
fn test_js_closure_captures_parameter_defaults() {
    let src = r#"
function fn(a = 10, getA = () => a) {
    a = 20;
    return getA();
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_closure_private_data_hiding_pattern() {
    let src = r#"
function BankAccount(initialBalance) {
    let balance = initialBalance;
    this.deposit = (amt) => { balance += amt; };
    this.getBalance = () => balance;
}
const acc = new BankAccount(100);
acc.deposit(50);
console.log(acc.getBalance() + "|hasBalanceProp=" + ("balance" in acc));
"#;
    assert_eq!(run_js(src), vec!["150|hasBalanceProp=false"]);
}

#[test]
fn test_js_closure_delayed_execution_captures_latest_value() {
    let src = r#"
let value = "before";
const getVal = () => value;
value = "after";
console.log(getVal());
"#;
    assert_eq!(run_js(src), vec!["after"]);
}

#[test]
fn test_js_closure_in_object_methods() {
    let src = r#"
const module = (() => {
    let internalState = 0;
    return {
        increment() { internalState++; },
        read() { return internalState; }
    };
})();
module.increment();
module.increment();
console.log(module.read());
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_closure_eval_access_outer_closure_vars() {
    let src = r#"
function outer() {
    const hidden = 999;
    return () => eval("hidden");
}
console.log(outer()());
"#;
    assert_eq!(run_js(src), vec!["999"]);
}

#[test]
fn test_js_closure_try_catch_block_scope_capture() {
    let src = r#"
let getError;
try {
    throw new Error("CaughtInTry");
} catch (err) {
    getError = () => err.message;
}
console.log(getError());
"#;
    assert_eq!(run_js(src), vec!["CaughtInTry"]);
}

#[test]
fn test_js_closure_recursion_with_outer_environment() {
    let src = r#"
function makeFactorialMemo() {
    const memo = {};
    function fact(n) {
        if (n <= 1) return 1;
        if (memo[n]) return memo[n];
        return (memo[n] = n * fact(n - 1));
    }
    return fact;
}
const fact = makeFactorialMemo();
console.log(fact(5) + "|" + fact(5));
"#;
    assert_eq!(run_js(src), vec!["120|120"]);
}

#[test]
fn test_js_closure_with_destructured_parameters() {
    let src = r#"
function process({ name, age }) {
    return () => `${name} is ${age}`;
}
const desc = process({ name: "Bob", age: 25 });
console.log(desc());
"#;
    assert_eq!(run_js(src), vec!["Bob is 25"]);
}

#[test]
fn test_js_closure_currying_pattern() {
    let src = r#"
const curry3 = (fn) => (a) => (b) => (c) => fn(a, b, c);
const sum = (a, b, c) => a + b + c;
console.log(curry3(sum)(1)(2)(3));
"#;
    assert_eq!(run_js(src), vec!["6"]);
}

#[test]
fn test_js_closure_class_constructor_private_scope() {
    let src = r#"
class Widget {
    constructor(id) {
        this.getId = () => id;
    }
}
const w = new Widget("W123");
console.log(w.getId() + "|hasId=" + ("id" in w));
"#;
    assert_eq!(run_js(src), vec!["W123|hasId=false"]);
}

#[test]
fn test_js_closure_async_function_outer_capture() {
    let src = r#"
function asyncFetcher(url) {
    return async () => {
        return await Promise.resolve("Fetched: " + url);
    };
}
(async () => {
    const fetcher = asyncFetcher("https://api.com");
    console.log(await fetcher());
})();
"#;
    assert_eq!(run_js(src), vec!["Fetched: https://api.com"]);
}

#[test]
fn test_js_closure_generator_outer_capture() {
    let src = r#"
function makeSeqGenerator(step) {
    return function*() {
        let current = 0;
        while (true) {
            yield (current += step);
        }
    };
}
const gen = makeSeqGenerator(5)();
console.log(`${gen.next().value}:${gen.next().value}:${gen.next().value}`);
"#;
    assert_eq!(run_js(src), vec!["5:10:15"]);
}

#[test]
fn test_js_closure_shadowing_outer_variable() {
    let src = r#"
const x = "outer";
function outerFn() {
    const x = "inner";
    return () => x;
}
console.log(outerFn()());
"#;
    assert_eq!(run_js(src), vec!["inner"]);
}
