// vybe-test: js/numeric_coercion_deep/number_from_various_types
// origin: languages/js/tests/js/test_numeric_coercion_deep.rs

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

__check(__line(Number("42")), "42");
__check(__line(Number("  3.14  ")), "3.14"); // trims whitespace
__check(__line(Number("")), "0");
__check(__line(Number(null)), "0");
__check(__line(Number(undefined)), "NaN");
__check(__line(Number(true)), "1");
__check(__line(Number(false)), "0");
