// vybe-test: js/iterator_helpers_deep/iterator_map_helper
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

function* range(n) { for (let i = 0; i < n; i++) yield i; }
const it = range(5);
// Polyfill using generator
function* mapIter(iter, fn) {
    for (const v of iter) yield fn(v);
}
const doubled = [...mapIter(range(5), x => x * 2)];
console.log(doubled.join(","));
