// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_bigint64array_view_buffer
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

const b64 = new BigInt64Array([1000n]);
const clone = structuredClone(b64, { transfer: [b64.buffer] });
__check(__line((b64.buffer.byteLength === 0) + "|" + clone[0].toString()), "true|1000");
