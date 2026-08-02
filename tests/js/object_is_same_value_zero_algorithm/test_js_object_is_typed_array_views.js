// vybe-test: js/object_is_same_value_zero_algorithm/test_js_object_is_typed_array_views
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

const buf = new ArrayBuffer(16);
const v1 = new Int32Array(buf);
const v2 = new Int32Array(buf);
__check(__line(Object.is(v1, v1)), "true");
__check(__line(Object.is(v1, v2)), "false");
