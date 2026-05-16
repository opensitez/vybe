/// JavaScript generators: function*, yield, yield*, iterator protocol,
/// custom iterables, Symbol.iterator, infinite sequences, generator composition.

use super::helpers::run_js;

// ===================================================================
// BASIC GENERATORS
// ===================================================================

#[test] fn generator_basic_yield() {
    assert_eq!(run_js(r#"
function* gen() {
    yield 1;
    yield 2;
    yield 3;
}
let g = gen();
console.log(g.next().value);
console.log(g.next().value);
console.log(g.next().value);
console.log(g.next().done);
"#), &["1", "2", "3", "true"]);
}

#[test] fn generator_for_of() {
    assert_eq!(run_js(r#"
function* range(start, end) {
    for (let i = start; i <= end; i++) {
        yield i;
    }
}
for (let n of range(1, 5)) {
    console.log(n);
}
"#), &["1", "2", "3", "4", "5"]);
}

#[test] fn generator_with_return() {
    assert_eq!(run_js(r#"
function* gen() {
    yield 1;
    return 99;
    yield 2;
}
let g = gen();
console.log(g.next().value);
let r = g.next();
console.log(r.value);
console.log(r.done);
"#), &["1", "99", "true"]);
}

#[test] fn generator_early_return() {
    assert_eq!(run_js(r#"
function* gen() {
    yield 1;
    yield 2;
    yield 3;
}
let g = gen();
console.log(g.next().value);
console.log(g.return("stopped").value);
console.log(g.next().done);
"#), &["1", "stopped", "true"]);
}

#[test] fn generator_return_runs_finally() {
    assert_eq!(run_js(r#"
function* gen() {
    try {
        yield 1;
        yield 2;
    } finally {
        console.log("cleanup");
    }
}
let g = gen();
console.log(g.next().value);
let result = g.return("stopped");
console.log(result.value);
console.log(result.done);
"#), &["1", "cleanup", "stopped", "true"]);
}

#[test] fn generator_return_before_first_yield() {
    assert_eq!(run_js(r#"
function* gen() {
    yield 1;
    return 99;
}
let g = gen();
let result = g.return("stopped");
console.log(result.value);
console.log(result.done);
console.log(g.next().done);
"#), &["stopped", "true", "true"]);
}

#[test] fn generator_throw_before_first_yield_caught() {
    assert_eq!(run_js(r#"
function* guarded() {
    try {
        yield "ready";
    } catch (err) {
        console.log("caught: " + err.message);
        yield "handled";
    }
}
let g = guarded();
let result = g.throw(new Error("stop"));
console.log(result.value);
console.log(result.done);
"#), &["caught: stop", "handled", "false"]);
}

#[test] fn generator_rest_args_survive_fresh_throw() {
    assert_eq!(run_js(r#"
function* guarded(head, ...rest) {
    try {
        yield rest.length;
    } catch (err) {
        console.log(rest.join(","));
        yield err.message;
    }
}
let g = guarded("a", "b", "c");
let result = g.throw(new Error("stop"));
console.log(result.value);
console.log(result.done);
"#), &["b,c", "stop", "false"]);
}

#[test] fn generator_yield_receive_value() {
    assert_eq!(run_js(r#"
function* echo() {
    let msg = yield "ready";
    console.log("received: " + msg);
    yield "done";
}
let g = echo();
console.log(g.next().value);
console.log(g.next("hello").value);
"#), &["ready", "received: hello", "done"]);
}

#[test] fn generator_throw_caught_in_generator() {
    assert_eq!(run_js(r#"
function* guarded() {
    try {
        yield "ready";
        yield "after";
    } catch (err) {
        console.log("caught: " + err.message);
    }
}
let g = guarded();
console.log(g.next().value);
let result = g.throw(new Error("stop"));
console.log(result.done);
"#), &["ready", "caught: stop", "true"]);
}

#[test] fn generator_infinite_sequence() {
    assert_eq!(run_js(r#"
function* naturals() {
    let n = 1;
    while (true) {
        yield n++;
    }
}
let gen = naturals();
let results = [];
for (let i = 0; i < 5; i++) {
    results.push(gen.next().value);
}
console.log(results.join(","));
"#), &["1,2,3,4,5"]);
}

#[test] fn generator_fibonacci() {
    assert_eq!(run_js(r#"
function* fib() {
    let a = 0, b = 1;
    while (true) {
        yield a;
        [a, b] = [b, a + b];
    }
}
let g = fib();
let results = [];
for (let i = 0; i < 8; i++) {
    results.push(g.next().value);
}
console.log(results.join(","));
"#), &["0,1,1,2,3,5,8,13"]);
}

#[test] fn generator_yield_star() {
    assert_eq!(run_js(r#"
function* inner() {
    yield 2;
    yield 3;
}
function* outer() {
    yield 1;
    yield* inner();
    yield 4;
}
let results = [];
for (let v of outer()) results.push(v);
console.log(results.join(","));
"#), &["1,2,3,4"]);
}

// ===================================================================
// CUSTOM ITERABLES
// ===================================================================

#[test] fn custom_iterable_symbol_iterator() {
    assert_eq!(run_js(r#"
let range = {
    from: 1,
    to: 5,
    [Symbol.iterator]() {
        let current = this.from;
        let last = this.to;
        return {
            next() {
                if (current <= last) {
                    return { value: current++, done: false };
                }
                return { done: true };
            }
        };
    }
};
let result = [];
for (let n of range) result.push(n);
console.log(result.join(","));
"#), &["1,2,3,4,5"]);
}

#[test] fn spread_on_custom_iterable() {
    assert_eq!(run_js(r#"
function* gen() {
    yield 10;
    yield 20;
    yield 30;
}
let arr = [...gen()];
console.log(arr.join(","));
"#), &["10,20,30"]);
}

#[test] fn destructure_generator() {
    assert_eq!(run_js(r#"
function* gen() {
    yield "a";
    yield "b";
    yield "c";
}
let [x, y, z] = gen();
console.log(x);
console.log(y);
console.log(z);
"#), &["a", "b", "c"]);
}

// ===================================================================
// ITERATOR PROTOCOL MANUAL
// ===================================================================

#[test] fn manual_iterator() {
    assert_eq!(run_js(r#"
let iter = {
    items: ["x", "y", "z"],
    index: 0,
    next() {
        if (this.index < this.items.length) {
            return { value: this.items[this.index++], done: false };
        }
        return { done: true };
    },
    [Symbol.iterator]() { return this; }
};
for (let v of iter) console.log(v);
"#), &["x", "y", "z"]);
}

#[test] fn array_from_generator() {
    assert_eq!(run_js(r#"
function* countdown(n) {
    while (n > 0) yield n--;
}
let arr = Array.from(countdown(5));
console.log(arr.join(","));
"#), &["5,4,3,2,1"]);
}
