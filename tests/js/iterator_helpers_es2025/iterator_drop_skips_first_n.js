// vybe-test: js/iterator_helpers_es2025/iterator_drop_skips_first_n
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

const result = Iterator.from([1, 2, 3, 4, 5]).drop(2).toArray();
__check(__line(result.join(",")), "3,4,5");
