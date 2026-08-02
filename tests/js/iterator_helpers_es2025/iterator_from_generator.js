// vybe-test: js/iterator_helpers_es2025/iterator_from_generator
// origin: languages/js/tests/js/test_iterator_helpers_es2025.rs

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

function* gen() { yield 1; yield 2; yield 3; }
const result = Iterator.from(gen()).toArray();
__check(__line(result.join(",")), "1,2,3");
