// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_uint8_overflow_wrap
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

const u8 = new Uint8Array(2);
u8[0] = 255;
u8[1] = 256; // Wraps modulo 256
__check(__line(`${u8[0]}:${u8[1]}`), "255:0");
