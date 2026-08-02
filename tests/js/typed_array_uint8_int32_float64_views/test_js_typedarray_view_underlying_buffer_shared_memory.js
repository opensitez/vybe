// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_view_underlying_buffer_shared_memory
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

const buf = new ArrayBuffer(4);
const u8 = new Uint8Array(buf);
const u32 = new Uint32Array(buf);

u8[0] = 0xFF;
__check(__line(u32[0] !== 0), "true"); // Modifying u8 updates shared u32 view in buffer
