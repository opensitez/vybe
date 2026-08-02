// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_empty_array
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

const buf = new Uint8Array([5]).buffer;
const clone = structuredClone(buf, { transfer: [] });
__check(__line((buf.byteLength === 1) + "|" + (clone.byteLength === 1)), "true|true");
