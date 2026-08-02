// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_constructor_with_array_like
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

const arr = new Int16Array([100, 200, 300]);
__check(__line(arr.length + "|" + arr[1]), "3|200");
