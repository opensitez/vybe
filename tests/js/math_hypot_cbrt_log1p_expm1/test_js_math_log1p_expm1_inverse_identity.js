// vybe-test: js/math_hypot_cbrt_log1p_expm1/test_js_math_log1p_expm1_inverse_identity
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

const x = 0.5;
const restored = Math.log1p(Math.expm1(x));
__check(__line(restored.toFixed(1)), "0.5");
