// vybe-test: js/type_coercion_deep/comparison_with_null_and_zero
// origin: languages/js/tests/js/test_type_coercion_deep.rs

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

// These are counterintuitive
__check(__line(null > 0), "false");
__check(__line(null == 0), "false");
__check(__line(null >= 0), "true");
