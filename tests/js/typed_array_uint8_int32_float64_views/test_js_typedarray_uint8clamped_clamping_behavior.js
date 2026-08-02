// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_uint8clamped_clamping_behavior
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

const clamped = new Uint8ClampedArray(3);
clamped[0] = -10; // Clamped to 0
clamped[1] = 300; // Clamped to 255
clamped[2] = 2.5; // Rounding half to even: 2
__check(__line(clamped.join(",")), "0,255,2");
