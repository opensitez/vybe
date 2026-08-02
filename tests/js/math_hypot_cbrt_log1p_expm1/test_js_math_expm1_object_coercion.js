// vybe-test: js/math_hypot_cbrt_log1p_expm1/test_js_math_expm1_object_coercion
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

const obj = { valueOf: () => 0 };
__check(__line(Math.expm1(obj)), "0");
