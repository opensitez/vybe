// vybe-test: js/iterator_patterns_deep/iterator_to_array_methods
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

function* range(start, end) {
    for (let i = start; i <= end; i++) yield i;
}
// Array.from accepts iterables
const arr = Array.from(range(1, 5));
console.log(arr.join(","));
// Spread also works
const doubled = [...range(1, 3)].map(x => x * 2);
console.log(doubled.join(","));
