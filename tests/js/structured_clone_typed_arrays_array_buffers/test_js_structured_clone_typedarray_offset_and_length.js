// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_typedarray_offset_and_length
// origin: languages/js/tests/js/test_js_structured_clone_typed_arrays_array_buffers.rs

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

const u8Base = new Uint8Array([0, 10, 20, 30, 0]);
const u8Sub = new Uint8Array(u8Base.buffer, 1, 3);
const clone = structuredClone(u8Sub);
__check(__line(clone.length + "|" + clone.byteOffset + "|" + clone.join(",")), "3|0|10,20,30");
