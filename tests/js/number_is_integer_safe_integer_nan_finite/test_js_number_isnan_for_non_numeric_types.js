// vybe-test: js/number_is_integer_safe_integer_nan_finite/test_js_number_isnan_for_non_numeric_types
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

__check(__line(`${Number.isNaN(undefined)}:${Number.isNaN({})}:${Number.isNaN("hello")}:${Number.isNaN(true)}`), "false:false:false:false");
