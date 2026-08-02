// vybe-test: js/iterator_helpers_deep/iterator_to_array_pattern
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

// Polyfill for Iterator.prototype.toArray
function toArray(iter) { return [...iter]; }
function* gen() { yield 1; yield 2; yield 3; }
console.log(toArray(gen()).join(","));
