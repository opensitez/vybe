// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_int32_signed_overflow
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

const i32 = new Int32Array(1);
i32[0] = 2147483647 + 1; // Signed 32-bit wrap
__check(__line(i32[0]), "-2147483648");
