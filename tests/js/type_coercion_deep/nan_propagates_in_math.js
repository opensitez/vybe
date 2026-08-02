// vybe-test: js/type_coercion_deep/nan_propagates_in_math
// origin: languages/js/tests/js/test_type_coercion_deep.rs

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

__check(__line(NaN + 1), "NaN");
__check(__line(NaN * 5), "NaN");
__check(__line(NaN - NaN), "NaN");
__check(__line(0 / 0), "NaN");
