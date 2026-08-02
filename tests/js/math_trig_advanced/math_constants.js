// vybe-test: js/math_trig_advanced/math_constants
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

__check(__line(typeof Math.PI), "number");
__check(__line(typeof Math.E), "number");
__check(__line(typeof Math.LN2), "number");
__check(__line(typeof Math.LN10), "number");
__check(__line(typeof Math.LOG2E), "number");
__check(__line(typeof Math.LOG10E), "number");
__check(__line(typeof Math.SQRT2), "number");
__check(__line(typeof Math.SQRT1_2), "number");
