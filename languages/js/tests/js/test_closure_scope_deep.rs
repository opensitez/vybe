/// Closure and scope deep — module pattern, private state, closure factories,
/// memoization patterns, event emitters, currying advanced, partial application,
/// function composition, trampolines, mutual recursion.
use super::helpers::run_js;

// ── private state via closure ─────────────────────────────────────────────────

#[test]
fn counter_via_closure() {
    assert_eq!(
        run_js(
            r#"
function makeCounter(start = 0) {
    let count = start;
    return {
        increment() { return ++count; },
        decrement() { return --count; },
        reset() { count = start; return count; },
        value() { return count; }
    };
}
const c = makeCounter(10);
console.log(c.increment());
console.log(c.increment());
console.log(c.decrement());
console.log(c.reset());
"#
        ),
        vec!["11", "12", "11", "10"]
    );
}

#[test]
fn multiple_closures_independent_state() {
    assert_eq!(
        run_js(
            r#"
function makeAdder(n) {
    return x => x + n;
}
const add5 = makeAdder(5);
const add10 = makeAdder(10);
console.log(add5(3));
console.log(add10(3));
console.log(add5(10) === add10(5));
"#
        ),
        vec!["8", "13", "true"]
    );
}

// ── event emitter pattern ─────────────────────────────────────────────────────

#[test]
fn event_emitter_via_closure() {
    assert_eq!(
        run_js(
            r#"
function createEmitter() {
    const listeners = new Map();
    return {
        on(event, fn) {
            if (!listeners.has(event)) listeners.set(event, []);
            listeners.get(event).push(fn);
        },
        emit(event, data) {
            listeners.get(event)?.forEach(fn => fn(data));
        }
    };
}
const em = createEmitter();
const log = [];
em.on("data", v => log.push("a:" + v));
em.on("data", v => log.push("b:" + v));
em.emit("data", 42);
console.log(log.join(","));
"#
        ),
        vec!["a:42,b:42"]
    );
}

// ── memoization ───────────────────────────────────────────────────────────────

#[test]
fn memoize_with_cache_clear() {
    assert_eq!(
        run_js(
            r#"
function memoize(fn) {
    const cache = new Map();
    const memo = function(...args) {
        const key = JSON.stringify(args);
        if (!cache.has(key)) cache.set(key, fn(...args));
        return cache.get(key);
    };
    memo.clear = () => cache.clear();
    memo.size = () => cache.size;
    return memo;
}

let calls = 0;
const expensive = memoize((n) => { calls++; return n * n; });
expensive(5); expensive(5); expensive(6);
console.log(calls);          // 2 unique calls
console.log(expensive.size()); // 2 cached
expensive.clear();
console.log(expensive.size()); // 0
"#
        ),
        vec!["2", "2", "0"]
    );
}

// ── currying ──────────────────────────────────────────────────────────────────

#[test]
fn curry_auto_curried_function() {
    assert_eq!(
        run_js(
            r#"
function curry(fn) {
    return function curried(...args) {
        if (args.length >= fn.length) return fn(...args);
        return (...more) => curried(...args, ...more);
    };
}

const add = curry((a, b, c) => a + b + c);
console.log(add(1)(2)(3));
console.log(add(1, 2)(3));
console.log(add(1)(2, 3));
console.log(add(1, 2, 3));
"#
        ),
        vec!["6", "6", "6", "6"]
    );
}

// ── function composition ──────────────────────────────────────────────────────

#[test]
fn compose_right_to_left() {
    assert_eq!(
        run_js(
            r#"
const compose = (...fns) => x => fns.reduceRight((v, f) => f(v), x);
const double = x => x * 2;
const addOne = x => x + 1;
const square = x => x * x;

const transform = compose(double, addOne, square);
// square(3) = 9, addOne(9) = 10, double(10) = 20
console.log(transform(3));
"#
        ),
        vec!["20"]
    );
}

#[test]
fn pipe_left_to_right() {
    assert_eq!(
        run_js(
            r#"
const pipe = (...fns) => x => fns.reduce((v, f) => f(v), x);
const process = pipe(
    s => s.trim(),
    s => s.toLowerCase(),
    s => s.replace(/\s+/g, "-")
);
console.log(process("  Hello World  "));
"#
        ),
        vec!["hello-world"]
    );
}

// ── trampoline ────────────────────────────────────────────────────────────────

#[test]
fn trampoline_for_stack_safe_recursion() {
    assert_eq!(
        run_js(
            r#"
function trampoline(fn) {
    return function(...args) {
        let result = fn(...args);
        while (typeof result === "function") result = result();
        return result;
    };
}

// Stack-safe sum via trampoline
function sum(n, acc = 0) {
    if (n === 0) return acc;
    return () => sum(n - 1, acc + n);
}

const safeSum = trampoline(sum);
console.log(safeSum(100));
"#
        ),
        vec!["5050"]
    );
}

// ── partial application ───────────────────────────────────────────────────────

#[test]
fn partial_application_binds_leading_args() {
    assert_eq!(
        run_js(
            r#"
function partial(fn, ...preset) {
    return (...rest) => fn(...preset, ...rest);
}

const multiply = (a, b) => a * b;
const triple = partial(multiply, 3);
console.log(triple(5));
console.log(triple(10));
"#
        ),
        vec!["15", "30"]
    );
}

// ── mutual recursion ──────────────────────────────────────────────────────────

#[test]
fn mutual_recursion_even_odd() {
    assert_eq!(
        run_js(
            r#"
function isEven(n) {
    if (n === 0) return true;
    return isOdd(n - 1);
}
function isOdd(n) {
    if (n === 0) return false;
    return isEven(n - 1);
}
console.log(isEven(4));
console.log(isOdd(7));
console.log(isEven(1));
"#
        ),
        vec!["true", "true", "false"]
    );
}

// ── closure in async context ──────────────────────────────────────────────────

#[test]
fn closure_over_async_state() {
    assert_eq!(
        run_js(
            r#"
function createAsyncCounter() {
    let count = 0;
    return async function() {
        await Promise.resolve(); // yield control
        return ++count;
    };
}
const next = createAsyncCounter();
Promise.all([next(), next(), next()]).then(results => {
    console.log(results.join(","));
});
"#
        ),
        vec!["1,2,3"]
    );
}

// ── object factory with closure ───────────────────────────────────────────────

#[test]
fn factory_with_private_state_and_methods() {
    assert_eq!(
        run_js(
            r#"
function createStack() {
    const items = [];
    return {
        push(v) { items.push(v); return this; },
        pop() { return items.pop(); },
        peek() { return items[items.length - 1]; },
        size() { return items.length; },
        isEmpty() { return items.length === 0; }
    };
}

const s = createStack();
s.push(1).push(2).push(3);
console.log(s.size());
console.log(s.peek());
console.log(s.pop());
console.log(s.isEmpty());
"#
        ),
        vec!["3", "3", "3", "false"]
    );
}
