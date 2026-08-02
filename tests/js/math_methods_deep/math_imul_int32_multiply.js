// vybe-test: js/math_methods_deep/math_imul_int32_multiply
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

__check(__line(Math.imul(3, 4)), "12");
__check(__line(Math.imul(0xffffffff, 5)), "-5"); // int32 overflow
