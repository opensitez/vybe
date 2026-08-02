// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_unaligned_byte_offset_throws_rangeerror
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

const buf = new ArrayBuffer(16);
try {
    new Int32Array(buf, 3); // ByteOffset 3 is not a multiple of Int32 element size (4)!
} catch (e) {
    __check(__line("Unaligned ByteOffset RangeError"), "Unaligned ByteOffset RangeError");
}
