// vybe-test: js/object_is_same_value_zero_algorithm/test_js_same_value_zero_custom_polyfill_simulation
// origin: languages/js/tests/js/test_js_object_is_same_value_zero_algorithm.rs

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

function sameValueZero(x, y) {
    if (x === y) {
        return true; // Handles +0 and -0 returning true
    }
    return Number.isNaN(x) && Number.isNaN(y);
}
__check(__line(sameValueZero(NaN, NaN) + "|" + sameValueZero(+0, -0) + "|" + sameValueZero(1, 2)), "true|true|false");
