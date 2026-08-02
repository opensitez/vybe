// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_multiple_views_same_buffer_preserve_identity
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

const buf = new ArrayBuffer(8);
const view1 = new Uint8Array(buf);
const view2 = new Int32Array(buf);
const root = { v1: view1, v2: view2 };

const clone = structuredClone(root);
__check(__line(clone.v1.buffer === clone.v2.buffer), "true"); // Underlying buffer identity preserved!
