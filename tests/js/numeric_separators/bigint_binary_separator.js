// vybe-test: js/numeric_separators/bigint_binary_separator
// origin: languages/js/tests/js/test_numeric_separators.rs

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

const n = 0b1010_0101n;
__check(__line(n), "165n");
