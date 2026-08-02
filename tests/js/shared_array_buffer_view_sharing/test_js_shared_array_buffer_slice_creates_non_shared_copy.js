// vybe-test: js/shared_array_buffer_view_sharing/test_js_shared_array_buffer_slice_creates_non_shared_copy
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

const sab = new SharedArrayBuffer(8);
const u8 = new Uint8Array(sab);
u8[0] = 42;

const slicedSab = sab.slice(0, 4);
const slicedU8 = new Uint8Array(slicedSab);
slicedU8[0] = 99;

__check(__line(u8[0] + "|" + slicedU8[0]), "42|99");
