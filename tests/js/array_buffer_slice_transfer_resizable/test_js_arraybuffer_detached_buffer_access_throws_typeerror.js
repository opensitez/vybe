// vybe-test: js/array_buffer_slice_transfer_resizable/test_js_arraybuffer_detached_buffer_access_throws_typeerror
// origin: languages/js/tests/js/test_js_array_buffer_slice_transfer_resizable.rs

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
const u8 = new Uint8Array(buf);
buf.transfer(); // Detaches buf

try {
    u8[0] = 10;
} catch (e) {
    __check(__line("Detached Buffer Access TypeError"), "Detached Buffer Access TypeError");
}
