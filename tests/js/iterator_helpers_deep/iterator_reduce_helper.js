// vybe-test: js/iterator_helpers_deep/iterator_reduce_helper
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

function reduceIter(iter, fn, init) {
    let acc = init;
    for (const v of iter) acc = fn(acc, v);
    return acc;
}
function* range(start, end) { for (let i = start; i < end; i++) yield i; }
const sum = reduceIter(range(1, 6), (a, b) => a + b, 0);
console.log(sum);
