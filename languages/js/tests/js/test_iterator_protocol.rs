/// Iterator protocol — custom Symbol.iterator, return method, infinite iterators,
/// iterator adapters, generator as iterator, spread with custom iterables,
/// destructuring with custom iterables, iterator composition.
use super::helpers::run_js;

// ── basic Symbol.iterator ─────────────────────────────────────────────────────

#[test]
fn custom_iterable_with_for_of() {
    assert_eq!(
        run_js(
            r#"
const iterable = {
    [Symbol.iterator]() {
        let i = 0;
        return {
            next() {
                return i < 3
                    ? { value: i++, done: false }
                    : { value: undefined, done: true };
            }
        };
    }
};
const results = [];
for (const v of iterable) results.push(v);
console.log(results.join(","));
"#
        ),
        vec!["0,1,2"]
    );
}

#[test]
fn custom_iterable_spread() {
    assert_eq!(
        run_js(
            r#"
const iterable = {
    [Symbol.iterator]() {
        const vals = [10, 20, 30];
        let i = 0;
        return { next() { return i < vals.length ? { value: vals[i++], done: false } : { done: true }; } };
    }
};
const arr = [...iterable];
console.log(arr.join(","));
"#
        ),
        vec!["10,20,30"]
    );
}

#[test]
fn custom_iterable_destructuring() {
    assert_eq!(
        run_js(
            r#"
const range = {
    [Symbol.iterator]() {
        let n = 1;
        return { next() { return n <= 5 ? { value: n++, done: false } : { done: true }; } };
    }
};
const [a, b, c] = range;
console.log(a);
console.log(b);
console.log(c);
"#
        ),
        vec!["1", "2", "3"]
    );
}

#[test]
fn custom_iterable_array_from() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    [Symbol.iterator]() {
        let i = 0;
        return { next() { return i < 4 ? { value: i * i, done: false } : { done: true }; i++ } };
    }
};
// Simpler approach
function* gen() { for (let i = 0; i < 4; i++) yield i * i; }
const arr = Array.from(gen());
console.log(arr.join(","));
"#
        ),
        vec!["0,1,4,9"]
    );
}

// ── iterator return method ────────────────────────────────────────────────────

#[test]
fn iterator_return_called_on_break() {
    assert_eq!(
        run_js(
            r#"
const log = [];
const iterable = {
    [Symbol.iterator]() {
        let i = 0;
        return {
            next() { return { value: i++, done: false }; },
            return() { log.push("return called"); return { done: true }; }
        };
    }
};
for (const v of iterable) {
    if (v >= 2) break;
}
console.log(log.join(","));
"#
        ),
        vec!["return called"]
    );
}

#[test]
fn iterator_return_called_on_throw() {
    assert_eq!(
        run_js(
            r#"
const log = [];
const iterable = {
    [Symbol.iterator]() {
        let i = 0;
        return {
            next() { return { value: i++, done: false }; },
            return() { log.push("cleanup"); return { done: true }; }
        };
    }
};
try {
    for (const v of iterable) {
        if (v === 1) throw new Error("stop");
    }
} catch {}
console.log(log.join(","));
"#
        ),
        vec!["cleanup"]
    );
}

// ── infinite iterators ────────────────────────────────────────────────────────

#[test]
fn infinite_iterator_with_take_via_break() {
    assert_eq!(
        run_js(
            r#"
function* naturals() {
    let n = 1;
    while (true) yield n++;
}
const first5 = [];
for (const n of naturals()) {
    if (n > 5) break;
    first5.push(n);
}
console.log(first5.join(","));
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

#[test]
fn fibonacci_generator_first_ten() {
    assert_eq!(
        run_js(
            r#"
function* fib() {
    let [a, b] = [0, 1];
    while (true) { yield a; [a, b] = [b, a + b]; }
}
const result = [];
for (const n of fib()) {
    result.push(n);
    if (result.length >= 10) break;
}
console.log(result.join(","));
"#
        ),
        vec!["0,1,1,2,3,5,8,13,21,34"]
    );
}

// ── iterator adapters ─────────────────────────────────────────────────────────

#[test]
fn map_adapter_over_custom_iterable() {
    assert_eq!(
        run_js(
            r#"
function* map(iter, fn) {
    for (const v of iter) yield fn(v);
}
function* range(n) {
    for (let i = 0; i < n; i++) yield i;
}
const result = [...map(range(4), x => x * x)];
console.log(result.join(","));
"#
        ),
        vec!["0,1,4,9"]
    );
}

#[test]
fn filter_adapter_over_iterable() {
    assert_eq!(
        run_js(
            r#"
function* filter(iter, pred) {
    for (const v of iter) if (pred(v)) yield v;
}
function* range(n) { for (let i = 0; i < n; i++) yield i; }
const evens = [...filter(range(8), n => n % 2 === 0)];
console.log(evens.join(","));
"#
        ),
        vec!["0,2,4,6"]
    );
}

#[test]
fn take_adapter_limits_iterator() {
    assert_eq!(
        run_js(
            r#"
function* take(iter, n) {
    let count = 0;
    for (const v of iter) {
        if (count++ >= n) break;
        yield v;
    }
}
function* naturals() { let n = 1; while (true) yield n++; }
console.log([...take(naturals(), 5)].join(","));
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

#[test]
fn zip_two_iterables() {
    assert_eq!(
        run_js(
            r#"
function* zip(a, b) {
    const ia = a[Symbol.iterator]();
    const ib = b[Symbol.iterator]();
    while (true) {
        const ra = ia.next(), rb = ib.next();
        if (ra.done || rb.done) break;
        yield [ra.value, rb.value];
    }
}
const pairs = [...zip([1, 2, 3], ["a", "b", "c"])];
console.log(pairs.map(([n, l]) => n + l).join(","));
"#
        ),
        vec!["1a,2b,3c"]
    );
}

// ── object as iterable ────────────────────────────────────────────────────────

#[test]
fn object_is_self_iterating_via_symbol_iterator() {
    assert_eq!(
        run_js(
            r#"
class Range {
    constructor(start, end) { this.start = start; this.end = end; }
    [Symbol.iterator]() {
        let cur = this.start;
        const end = this.end;
        return {
            next() {
                return cur <= end
                    ? { value: cur++, done: false }
                    : { done: true };
            }
        };
    }
}
console.log([...new Range(3, 6)].join(","));
"#
        ),
        vec!["3,4,5,6"]
    );
}

// ── iterator protocol completeness ────────────────────────────────────────────

#[test]
fn iterator_next_value_is_undefined_after_done() {
    assert_eq!(
        run_js(
            r#"
function* gen() { yield 1; }
const g = gen();
g.next();
const r = g.next();
console.log(r.done);
console.log(r.value);
"#
        ),
        vec!["true", "undefined"]
    );
}

#[test]
fn iterable_can_be_iterated_multiple_times() {
    assert_eq!(
        run_js(
            r#"
const iterable = {
    [Symbol.iterator]() {
        let i = 0;
        return { next() { return i < 3 ? { value: i++, done: false } : { done: true }; } };
    }
};
const r1 = [...iterable];
const r2 = [...iterable];
console.log(r1.join(","));
console.log(r2.join(","));
"#
        ),
        vec!["0,1,2", "0,1,2"]
    );
}

// ── flat-map iteration ────────────────────────────────────────────────────────

#[test]
fn flatmap_using_generator_delegation() {
    assert_eq!(
        run_js(
            r#"
function* flatMap(iter, fn) {
    for (const v of iter) yield* fn(v);
}
const result = [...flatMap([1, 2, 3], n => [n, n * 10])];
console.log(result.join(","));
"#
        ),
        vec!["1,10,2,20,3,30"]
    );
}
