// vybe-test: js/number_is_integer_safe_integer_nan_finite/test_js_number_isfinite_vs_global_isfinite
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

__check(__line(`${Number.isFinite(100)}:${Number.isFinite("100")}:${isFinite("100")}:${Number.isFinite(Infinity)}`), "true:false:true:false");
