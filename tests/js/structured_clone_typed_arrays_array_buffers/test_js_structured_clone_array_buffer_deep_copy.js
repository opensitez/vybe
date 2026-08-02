// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_array_buffer_deep_copy
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

const buf = new Uint8Array([10, 20, 30]).buffer;
const cloneBuf = structuredClone(buf);
const u8 = new Uint8Array(cloneBuf);
__check(__line((cloneBuf !== buf) + "|" + u8.join(",")), "true|10,20,30");
