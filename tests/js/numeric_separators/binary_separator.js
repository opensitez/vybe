// vybe-test: js/numeric_separators/binary_separator
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

const flags = 0b1010_0001;
__check(__line(flags), "161");
