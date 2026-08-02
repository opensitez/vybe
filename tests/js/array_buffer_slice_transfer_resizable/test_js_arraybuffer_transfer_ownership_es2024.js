// vybe-test: js/array_buffer_slice_transfer_resizable/test_js_arraybuffer_transfer_ownership_es2024
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
const view = new Uint8Array(buf);
view[0] = 42;

const transferred = buf.transfer();
__check(__line(transferred.byteLength + "|" + new Uint8Array(transferred)[0] + "|oldDetached=" + buf.detached), "16|42|oldDetached=true");
