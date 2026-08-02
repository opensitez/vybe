// vybe-test: js/iterator_patterns_deep/lazy_sequence_operations
// origin: languages/js/tests/js/test_iterator_patterns_deep.rs

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

function* seq(arr) { yield* arr; }
function* mapLazy(gen, fn) { for (const v of gen) yield fn(v); }
function reduce(gen, fn, init) {
    let acc = init;
    for (const v of gen) acc = fn(acc, v);
    return acc;
}
const sum = reduce(mapLazy(seq([1,2,3,4,5]), x => x*x), (a,b) => a+b, 0);
console.log(sum);
