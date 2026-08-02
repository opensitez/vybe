// vybe-test: js/math_methods_deep/math_abs_various
// origin: languages/js/tests/js/test_math_methods_deep.rs

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

__check(__line(Math.abs(-5)), "5");
__check(__line(Math.abs(5)), "5");
__check(__line(Math.abs(-Infinity)), "Infinity");
__check(__line(Math.abs(0)), "0");
