// vybe-test: js/numeric_separators/separator_in_computed_expression
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

const kb = 1_024;
const mb = 1_024 * 1_024;
__check(__line(mb / kb), "1024");
