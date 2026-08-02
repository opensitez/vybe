// vybe-test: js/string_array_advanced/number_is_finite
// origin: languages/js/tests/js/test_string_array_advanced.rs

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

__check(__line(Number.isFinite(42)), "true");
__check(__line(Number.isFinite(Infinity)), "false");
__check(__line(Number.isFinite(NaN)), "false");
