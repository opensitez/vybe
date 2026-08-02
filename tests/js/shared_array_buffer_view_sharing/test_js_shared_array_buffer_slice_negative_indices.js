// vybe-test: js/shared_array_buffer_view_sharing/test_js_shared_array_buffer_slice_negative_indices
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

const sab = new SharedArrayBuffer(10);
const u8 = new Uint8Array(sab);
u8[8] = 99;
const sliced = new Uint8Array(sab.slice(-2));
__check(__line(sliced.length + "|" + sliced[0]), "2|99");
