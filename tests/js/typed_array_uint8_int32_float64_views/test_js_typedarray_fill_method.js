// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_fill_method
// origin: languages/js/tests/js/test_js_typed_array_uint8_int32_float64_views.rs

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

const arr = new Uint8Array(4);
arr.fill(42, 1, 3);
__check(__line(arr.join(",")), "0,42,42,0");
