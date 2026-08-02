// vybe-test: js/math_methods_deep/math_clamp_pattern
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

function clamp(val, min, max) {
    return Math.min(Math.max(val, min), max);
}
__check(__line(clamp(5, 0, 10)), "5");
__check(__line(clamp(-5, 0, 10)), "0");
__check(__line(clamp(15, 0, 10)), "10");
