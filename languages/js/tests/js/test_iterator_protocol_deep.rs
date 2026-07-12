/// Iterator protocol deep — Symbol.iterator, return method, throw method, custom iterators
use super::helpers::run_js;

#[test]
fn custom_iterator_protocol() {
    assert_eq!(
        run_js(
            r#"
function makeCounter(max) {
    let n = 0;
    return {
        [Symbol.iterator]() { return this; },
        next() {
            return n < max ? { value: n++, done: false } : { done: true, value: undefined };
        }
    };
}
console.log([...makeCounter(3)].join(","));
"#
        ),
        vec!["0,1,2"]
    );
}

#[test]
fn iterator_return_called_on_break() {
    assert_eq!(
        run_js(
            r#"
let returnCalled = false;
const iterable = {
    [Symbol.iterator]() {
        let n = 0;
        return {
            next() { return n < 10 ? { value: n++, done: false } : { done: true }; },
            return() { returnCalled = true; return { done: true }; }
        };
    }
};
for (const v of iterable) {
    if (v === 2) break;
}
console.log(returnCalled);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn iterator_return_called_on_throw() {
    assert_eq!(
        run_js(
            r#"
let returnCalled = false;
const iterable = {
    [Symbol.iterator]() {
        let n = 0;
        return {
            next() { return { value: n++, done: false }; },
            return() { returnCalled = true; return { done: true }; }
        };
    }
};
try {
    for (const v of iterable) {
        if (v === 2) throw new Error("stop");
    }
} catch {}
console.log(returnCalled);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn iterable_object_can_be_spread() {
    assert_eq!(
        run_js(
            r#"
const range = {
    from: 1, to: 5,
    [Symbol.iterator]() {
        let cur = this.from, end = this.to;
        return { next() { return cur <= end ? { value: cur++, done: false } : { done: true }; } };
    }
};
console.log([...range].join(","));
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

#[test]
fn destructuring_uses_iterator() {
    assert_eq!(
        run_js(
            r#"
function* gen() { yield 1; yield 2; yield 3; }
const [a, b, c] = gen();
console.log(a);
console.log(b);
console.log(c);
"#
        ),
        vec!["1", "2", "3"]
    );
}

#[test]
fn for_of_calls_next_until_done() {
    assert_eq!(
        run_js(
            r#"
let nextCalls = 0;
const it = {
    [Symbol.iterator]() { return this; },
    next() {
        nextCalls++;
        return nextCalls <= 3 ? { value: nextCalls, done: false } : { done: true };
    }
};
const vals = [];
for (const v of it) vals.push(v);
console.log(vals.join(","));
console.log(nextCalls); // includes the final done call
"#
        ),
        vec!["1,2,3", "4"]
    );
}

#[test]
fn string_is_iterable() {
    assert_eq!(
        run_js(
            r#"
const chars = [];
for (const c of "hello") chars.push(c);
console.log(chars.join("-"));
"#
        ),
        vec!["h-e-l-l-o"]
    );
}

#[test]
fn map_is_iterable_yields_entries() {
    assert_eq!(
        run_js(
            r#"
const m = new Map([["a", 1], ["b", 2]]);
const pairs = [];
for (const [k, v] of m) pairs.push(k + "=" + v);
console.log(pairs.join(","));
"#
        ),
        vec!["a=1,b=2"]
    );
}

#[test]
fn iterator_result_done_false_value() {
    assert_eq!(
        run_js(
            r#"
function* gen() { return "final"; }
const it = gen();
const r1 = it.next();
console.log(r1.done);
console.log(r1.value);
const r2 = it.next();
console.log(r2.done);
console.log(r2.value);
"#
        ),
        vec!["true", "final", "true", "undefined"]
    );
}

#[test]
fn generator_is_both_iterable_and_iterator() {
    assert_eq!(
        run_js(
            r#"
function* gen() { yield 1; yield 2; }
const it = gen();
console.log(it[Symbol.iterator]() === it); // same object
"#
        ),
        vec!["true"]
    );
}

#[test]
fn well_formed_iterator_always_returns_object() {
    assert_eq!(
        run_js(
            r#"
// next() must always return {value, done}
function* g() { yield 1; }
const it = g();
const r = it.next();
console.log(typeof r);
console.log("value" in r);
console.log("done" in r);
"#
        ),
        vec!["object", "true", "true"]
    );
}
