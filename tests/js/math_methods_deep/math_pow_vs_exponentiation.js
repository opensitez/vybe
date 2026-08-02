// vybe-test: js/math_methods_deep/math_pow_vs_exponentiation
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

__check(__line(Math.pow(2, 10)), "1024");
__check(__line(2 ** 10), "1024");
__check(__line(Math.pow(2, 0.5).toFixed(4)), "1.4142");
