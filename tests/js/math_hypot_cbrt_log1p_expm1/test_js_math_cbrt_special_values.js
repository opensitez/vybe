// vybe-test: js/math_hypot_cbrt_log1p_expm1/test_js_math_cbrt_special_values
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

__check(__line(`${Math.cbrt(-0)}:${Math.cbrt(Infinity)}:${Math.cbrt(-Infinity)}:${Math.cbrt(NaN)}`), "-0:Infinity:-Infinity:NaN");
