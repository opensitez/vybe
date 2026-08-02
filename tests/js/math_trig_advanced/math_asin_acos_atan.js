// vybe-test: js/math_trig_advanced/math_asin_acos_atan
// origin: languages/js/tests/js/test_math_trig_advanced.rs

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

__check(__line(Math.asin(1).toFixed(5)), "1.57080");
__check(__line(Math.acos(1).toFixed(5)), "0.00000");
__check(__line(Math.atan(1).toFixed(5)), "0.78540");
