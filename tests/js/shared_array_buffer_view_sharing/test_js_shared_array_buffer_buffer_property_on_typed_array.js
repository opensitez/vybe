// vybe-test: js/shared_array_buffer_view_sharing/test_js_shared_array_buffer_buffer_property_on_typed_array
// origin: languages/js/tests/js/test_js_shared_array_buffer_view_sharing.rs

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

const sab = new SharedArrayBuffer(16);
const u8 = new Uint8Array(sab);
__check(__line(u8.buffer === sab + "|" + u8.buffer.byteLength), "true|16");
