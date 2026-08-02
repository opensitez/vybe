// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_multiple_buffers
// origin: languages/js/tests/js/test_js_structured_clone_transferables_array_buffer.rs

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

const buf1 = new Uint8Array([1]).buffer;
const buf2 = new Uint8Array([2]).buffer;
const root = { b1: buf1, b2: buf2 };
const clone = structuredClone(root, { transfer: [buf1, buf2] });

__check(__line((buf1.byteLength === 0) + "|" + (buf2.byteLength === 0) + "|" + new Uint8Array(clone.b1)[0] + "|" + new Uint8Array(clone.b2)[0]), "true|true|1|2");
