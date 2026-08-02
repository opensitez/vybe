// vybe-test: js/object_is_same_value_zero_algorithm/test_js_object_is_dates
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

const d1 = new Date(2025, 1, 1);
const d2 = new Date(2025, 1, 1);
__check(__line(Object.is(d1, d1)), "true");
__check(__line(Object.is(d1, d2)), "false");
