// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_with_primitive_target
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

const buf = new Uint8Array([99]).buffer;
const cloneVal = structuredClone(100, { transfer: [buf] });
__check(__line(cloneVal + "|detached=" + (buf.byteLength === 0)), "100|detached=true");
