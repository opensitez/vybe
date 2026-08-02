// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_array_buffer_transfer
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

const buf = new Uint8Array([1, 2, 3, 4]).buffer;
const clone = structuredClone(buf, { transfer: [buf] });
const u8 = new Uint8Array(clone);

__check(__line((buf.byteLength === 0) + "|" + u8.join(",")), "true|1,2,3,4"); // buf is transferred (detached, length 0)!
