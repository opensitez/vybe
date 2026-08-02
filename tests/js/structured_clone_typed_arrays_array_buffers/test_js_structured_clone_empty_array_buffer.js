// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_empty_array_buffer
// origin: languages/js/tests/js/test_js_structured_clone_typed_arrays_array_buffers.rs

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

const emptyBuf = new ArrayBuffer(0);
const clone = structuredClone(emptyBuf);
__check(__line(clone.byteLength), "0");
