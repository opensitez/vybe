// vybe-test: js/array_buffer_slice_transfer_resizable/test_js_arraybuffer_is_view_static_utility
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
const dv = new DataView(buf);

__check(__line(`${ArrayBuffer.isView(u8)}|${ArrayBuffer.isView(dv)}|${ArrayBuffer.isView(buf)}|${ArrayBuffer.isView({})}`), "true|true|false|false");
