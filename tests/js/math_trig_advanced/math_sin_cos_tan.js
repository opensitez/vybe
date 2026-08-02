// vybe-test: js/math_trig_advanced/math_sin_cos_tan
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

__check(__line(Math.sin(0)), "0");
__check(__line(Math.cos(0)), "1");
__check(__line(Math.tan(0)), "0");
__check(__line(Math.sin(Math.PI / 2).toFixed(10)), "1.0000000000");
