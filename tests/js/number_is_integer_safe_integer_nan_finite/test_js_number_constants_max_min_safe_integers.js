// vybe-test: js/number_is_integer_safe_integer_nan_finite/test_js_number_constants_max_min_safe_integers
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

__check(__line(`${Number.MAX_SAFE_INTEGER}:${Number.MIN_SAFE_INTEGER}`), "9007199254740991:-9007199254740991");
