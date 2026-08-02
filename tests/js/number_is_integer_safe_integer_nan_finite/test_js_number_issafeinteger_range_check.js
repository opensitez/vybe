// vybe-test: js/number_is_integer_safe_integer_nan_finite/test_js_number_issafeinteger_range_check
// origin: languages/js/tests/js/test_js_number_is_integer_safe_integer_nan_finite.rs

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

const maxSafe = Number.MAX_SAFE_INTEGER;
__check(__line(`${Number.isSafeInteger(maxSafe)}:${Number.isSafeInteger(maxSafe + 1)}:${Number.isSafeInteger(-maxSafe)}:${Number.isSafeInteger(-maxSafe - 1)}`), "true:false:true:false");
