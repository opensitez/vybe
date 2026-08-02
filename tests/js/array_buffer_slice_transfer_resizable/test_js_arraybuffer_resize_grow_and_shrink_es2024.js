// vybe-test: js/array_buffer_slice_transfer_resizable/test_js_arraybuffer_resize_grow_and_shrink_es2024
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

const buf = new ArrayBuffer(8, { maxByteLength: 32 });
buf.resize(16);
__check(__line(buf.byteLength), "16");
buf.resize(4);
__check(__line(buf.byteLength), "4");
