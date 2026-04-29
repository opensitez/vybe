/// JavaScript closure patterns, functional programming: currying, memoization,
/// module pattern, partial application, compose, pipe, higher-order patterns,
/// trampolining, once/debounce patterns.

use super::helpers::run_js;

// ===================================================================
// CURRYING
// ===================================================================

#[test]
fn curry_basic() {
    assert_eq!(run_js(r#"
function curry(fn) {
    return function(a) {
        return function(b) {
            return fn(a, b);
        };
    };
}
let add = curry((a, b) => a + b);
console.log(add(3)(4));
"#), &["7"]);
}

#[test]
fn curry_reusable() {
    assert_eq!(run_js(r#"
function multiply(a) {
    return function(b) {
        return a * b;
    };
}
let double = multiply(2);
let triple = multiply(3);
console.log(double(5));
console.log(triple(5));
"#), &["10", "15"]);
}

#[test]
fn curry_string_formatter() {
    assert_eq!(run_js(r#"
function greet(greeting) {
    return function(name) {
        return greeting + ", " + name + "!";
    };
}
let hello = greet("Hello");
let hi = greet("Hi");
console.log(hello("Alice"));
console.log(hi("Bob"));
"#), &["Hello, Alice!", "Hi, Bob!"]);
}

// ===================================================================
// PARTIAL APPLICATION
// ===================================================================

#[test]
fn partial_application() {
    assert_eq!(run_js(r#"
function partial(fn, ...presets) {
    return function(...args) {
        return fn(...presets, ...args);
    };
}
function add3(a, b, c) { return a + b + c; }
let addTo10 = partial(add3, 3, 7);
console.log(addTo10(5));
console.log(addTo10(10));
"#), &["15", "20"]);
}

// ===================================================================
// COMPOSE AND PIPE
// ===================================================================

#[test]
fn compose_functions() {
    assert_eq!(run_js(r#"
function compose(...fns) {
    return function(x) {
        return fns.reduceRight((acc, fn) => fn(acc), x);
    };
}
let addOne = x => x + 1;
let double = x => x * 2;
let square = x => x * x;
let transform = compose(square, double, addOne);
console.log(transform(3));
"#), &["64"]);
}

#[test]
fn pipe_functions() {
    assert_eq!(run_js(r#"
function pipe(...fns) {
    return function(x) {
        return fns.reduce((acc, fn) => fn(acc), x);
    };
}
let addOne = x => x + 1;
let double = x => x * 2;
let transform = pipe(addOne, double, addOne);
console.log(transform(3));
"#), &["9"]);
}

// ===================================================================
// MEMOIZATION
// ===================================================================

#[test]
fn memoize_basic() {
    assert_eq!(run_js(r#"
function memoize(fn) {
    let cache = {};
    return function(n) {
        if (cache[n] !== undefined) return cache[n];
        cache[n] = fn(n);
        return cache[n];
    };
}
let callCount = 0;
let square = memoize(n => { callCount++; return n * n; });
console.log(square(4));
console.log(square(4));
console.log(square(5));
console.log(callCount);
"#), &["16", "16", "25", "2"]);
}

#[test]
fn memoize_fibonacci() {
    assert_eq!(run_js(r#"
function memoize(fn) {
    let cache = {};
    return function(n) {
        if (cache[n] !== undefined) return cache[n];
        cache[n] = fn(n);
        return cache[n];
    };
}
let fib = memoize(function(n) {
    if (n <= 1) return n;
    return fib(n - 1) + fib(n - 2);
});
console.log(fib(10));
console.log(fib(20));
"#), &["55", "6765"]);
}

// ===================================================================
// MODULE PATTERN
// ===================================================================

#[test]
fn module_pattern_iife() {
    assert_eq!(run_js(r#"
let counter = (function() {
    let count = 0;
    return {
        increment() { count++; },
        decrement() { count--; },
        getCount() { return count; }
    };
})();
counter.increment();
counter.increment();
counter.increment();
counter.decrement();
console.log(counter.getCount());
"#), &["2"]);
}

#[test]
fn module_pattern_private_state() {
    assert_eq!(run_js(r#"
let bank = (function() {
    let balance = 0;
    return {
        deposit(amt) { balance += amt; },
        withdraw(amt) {
            if (amt > balance) return false;
            balance -= amt;
            return true;
        },
        getBalance() { return balance; }
    };
})();
bank.deposit(100);
bank.deposit(50);
console.log(bank.withdraw(30));
console.log(bank.getBalance());
console.log(bank.withdraw(200));
console.log(bank.getBalance());
"#), &["true", "120", "false", "120"]);
}

// ===================================================================
// CLOSURE PATTERNS
// ===================================================================

#[test]
fn closure_factory() {
    assert_eq!(run_js(r#"
function makeCounter(start) {
    let count = start;
    return {
        next() { return count++; },
        reset() { count = start; }
    };
}
let c = makeCounter(10);
console.log(c.next());
console.log(c.next());
console.log(c.next());
c.reset();
console.log(c.next());
"#), &["10", "11", "12", "10"]);
}

#[test]
fn closure_over_let_in_loop() {
    assert_eq!(run_js(r#"
let funcs = [];
for (let i = 0; i < 5; i++) {
    funcs.push(() => i);
}
console.log(funcs[0]());
console.log(funcs[2]());
console.log(funcs[4]());
"#), &["0", "2", "4"]);
}

#[test]
fn closure_over_var_in_loop() {
    let handle = std::thread::Builder::new()
        .name("closure_over_var_in_loop".to_string())
        .stack_size(8 * 1024 * 1024)
        .spawn(|| {
            assert_eq!(run_js(r#"
var funcs = [];
for (var i = 0; i < 5; i++) {
    funcs.push((function(j) { return function() { return j; }; })(i));
}
console.log(funcs[0]());
console.log(funcs[2]());
console.log(funcs[4]());
"#), &["0", "2", "4"]);
        })
        .expect("failed to spawn test thread");

    handle.join().expect("closure_over_var_in_loop thread panicked");
}

#[test]
fn once_function() {
    assert_eq!(run_js(r#"
function once(fn) {
    let called = false;
    let result;
    return function(...args) {
        if (!called) {
            called = true;
            result = fn(...args);
        }
        return result;
    };
}
let init = once(() => { console.log("initialized"); return 42; });
console.log(init());
console.log(init());
"#), &["initialized", "42", "42"]);
}

// ===================================================================
// HIGHER-ORDER PATTERNS
// ===================================================================

#[test]
fn map_filter_reduce_chain() {
    assert_eq!(run_js(r#"
let data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
let result = data
    .filter(x => x % 2 === 0)
    .map(x => x * x)
    .reduce((acc, x) => acc + x, 0);
console.log(result);
"#), &["220"]);
}

#[test]
fn custom_flat_map() {
    assert_eq!(run_js(r#"
let sentences = ["Hello World", "Foo Bar Baz"];
let words = sentences.flatMap(s => s.split(" "));
console.log(words.join(","));
"#), &["Hello,World,Foo,Bar,Baz"]);
}

#[test]
fn reduce_group_by() {
    assert_eq!(run_js(r#"
let people = [
    { name: "Alice", dept: "eng" },
    { name: "Bob", dept: "sales" },
    { name: "Charlie", dept: "eng" },
    { name: "Diana", dept: "sales" },
    { name: "Eve", dept: "eng" }
];
let groups = people.reduce((acc, p) => {
    if (!acc[p.dept]) acc[p.dept] = [];
    acc[p.dept].push(p.name);
    return acc;
}, {});
console.log(groups.eng.length);
console.log(groups.sales.length);
"#), &["3", "2"]);
}

#[test]
fn reduce_to_frequency_map() {
    assert_eq!(run_js(r#"
let letters = "abracadabra".split("");
let freq = letters.reduce((acc, ch) => {
    acc[ch] = (acc[ch] || 0) + 1;
    return acc;
}, {});
console.log(freq.a);
console.log(freq.b);
console.log(freq.r);
"#), &["5", "2", "2"]);
}

// ===================================================================
// TRAMPOLINING (tail-call optimization pattern)
// ===================================================================

#[test]
fn trampoline() {
    assert_eq!(run_js(r#"
function trampoline(fn) {
    return function(...args) {
        let result = fn(...args);
        while (typeof result === "function") {
            result = result();
        }
        return result;
    };
}
function sumHelper(n, acc) {
    if (n === 0) return acc;
    return () => sumHelper(n - 1, acc + n);
}
let tSum = trampoline(sumHelper);
console.log(tSum(100, 0));
"#), &["5050"]);
}
