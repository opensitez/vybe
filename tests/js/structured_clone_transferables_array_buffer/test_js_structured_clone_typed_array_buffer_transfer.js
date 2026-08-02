// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_typed_array_buffer_transfer
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

const u8 = new Uint8Array([10, 20, 30]);
const clone = structuredClone(u8, { transfer: [u8.buffer] });

__check(__line((u8.buffer.byteLength === 0) + "|" + (u8.length === 0 || u8[0] === undefined) + "|" + clone.join(",")), "true|true|10,20,30");
