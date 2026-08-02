// vybe-test: js/array_buffer_slice_transfer_resizable/test_js_arraybuffer_slice_copy_range
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
const view1 = new Uint8Array(buf);
view1[4] = 99;

const slicedBuf = buf.slice(4, 8);
const view2 = new Uint8Array(slicedBuf);
__check(__line(slicedBuf.byteLength + "|" + view2[0] + "|isCopy=" + (slicedBuf !== buf)), "4|99|isCopy=true");
