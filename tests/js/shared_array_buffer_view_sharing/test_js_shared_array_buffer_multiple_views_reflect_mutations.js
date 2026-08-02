// vybe-test: js/shared_array_buffer_view_sharing/test_js_shared_array_buffer_multiple_views_reflect_mutations
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

const sab = new SharedArrayBuffer(4);
const u8 = new Uint8Array(sab);
const i32 = new Int32Array(sab);

u8[0] = 0xFF; // Set first byte to 255
__check(__line(i32[0] & 0xFF), "255");
