// vybe-test: js/numeric_coercion_deep/number_to_fixed_rounding
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

__check(__line((1.005).toFixed(2)), "1.00");
__check(__line((1.255).toFixed(2)), "1.25");
__check(__line((1.5).toFixed(0)), "2");
