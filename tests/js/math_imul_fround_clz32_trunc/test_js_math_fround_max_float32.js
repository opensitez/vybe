// vybe-test: js/math_imul_fround_clz32_trunc/test_js_math_fround_max_float32
// origin: languages/js/tests/js/test_js_math_imul_fround_clz32_trunc.rs

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

const maxF32 = 3.4028234663852886e+38;
__check(__line(Math.fround(maxF32) === maxF32), "true");
