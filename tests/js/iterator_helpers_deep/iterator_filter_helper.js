// vybe-test: js/iterator_helpers_deep/iterator_filter_helper
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

function* filterIter(iter, pred) {
    for (const v of iter) if (pred(v)) yield v;
}
function* range(n) { for (let i = 0; i < n; i++) yield i; }
const evens = [...filterIter(range(10), x => x % 2 === 0)];
console.log(evens.join(","));
