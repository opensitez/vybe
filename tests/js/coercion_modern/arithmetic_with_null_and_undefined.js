// vybe-test: js/coercion_modern/arithmetic_with_null_and_undefined
// origin: languages/js/tests/js/test_coercion_modern.rs

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

__check(__line(null + 1), "1");
__check(__line(undefined + 1), "NaN");
__check(__line(null == 0), "false");
