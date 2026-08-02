// vybe-test: js/numeric_separators/scientific_notation_separator
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

const n = 1_000e2;
__check(__line(n), "100000");
