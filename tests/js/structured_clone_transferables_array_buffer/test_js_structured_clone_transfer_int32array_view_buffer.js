// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_int32array_view_buffer
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

const i32 = new Int32Array([100, 200]);
const clone = structuredClone(i32, { transfer: [i32.buffer] });
__check(__line((i32.buffer.byteLength === 0) + "|" + clone.join(",")), "true|100,200");
