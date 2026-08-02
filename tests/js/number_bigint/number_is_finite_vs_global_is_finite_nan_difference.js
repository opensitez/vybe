// vybe-test: js/number_bigint/number_is_finite_vs_global_is_finite_nan_difference
// origin: languages/js/tests/js/test_number_bigint.rs

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

__check(__line(Number.isFinite(42)), "true");
__check(__line(Number.isFinite("42")), "false");
__check(__line(isFinite("42")), "true");
