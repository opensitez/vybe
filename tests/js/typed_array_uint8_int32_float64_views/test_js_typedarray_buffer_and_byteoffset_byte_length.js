// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_buffer_and_byteoffset_byte_length
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

const buffer = new ArrayBuffer(16);
const view = new Int32Array(buffer, 4, 2);
__check(__line(`${view.length}:${view.byteOffset}:${view.byteLength}`), "2:4:8");
