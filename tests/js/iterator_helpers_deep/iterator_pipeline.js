// vybe-test: js/iterator_helpers_deep/iterator_pipeline
// origin: languages/js/tests/js/test_iterator_helpers_deep.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

// Chain operations lazily
function* range(n) { for (let i = 1; i <= n; i++) yield i; }
function* map(fn, iter) { for (const v of iter) yield fn(v); }
function* filter(pred, iter) { for (const v of iter) if (pred(v)) yield v; }
function* take(n, iter) {
    let i = 0;
    for (const v of iter) { if (i++ >= n) break; yield v; }
}
const pipeline = take(3, filter(x => x % 2 === 0, map(x => x * x, range(20))));
console.log([...pipeline].join(","));
