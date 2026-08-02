// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_cannot_delete_indexed_properties
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

const arr = new Uint8Array([10]);
try {
    "use strict";
    delete arr[0];
} catch (e) {
    __check(__line("Delete TypedArray Index TypeError"), "Delete TypedArray Index TypeError");
}
__check(__line(arr[0]), "10");
