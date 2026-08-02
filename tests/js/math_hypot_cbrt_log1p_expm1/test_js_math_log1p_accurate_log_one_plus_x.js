// vybe-test: js/math_hypot_cbrt_log1p_expm1/test_js_math_log1p_accurate_log_one_plus_x
// origin: languages/js/tests/js/test_js_math_hypot_cbrt_log1p_expm1.rs

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

__check(__line(`${Math.log1p(0)}:${Math.log1p(-1)}:${Math.log1p(1e-15) > 0}`), "0:-Infinity:true");
