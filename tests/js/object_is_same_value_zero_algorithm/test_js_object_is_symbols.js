// vybe-test: js/object_is_same_value_zero_algorithm/test_js_object_is_symbols
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

const s1 = Symbol("id");
const s2 = Symbol("id");
__check(__line(Object.is(s1, s1)), "true");
__check(__line(Object.is(s1, s2)), "false");
