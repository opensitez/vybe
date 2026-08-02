// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_resizable_array_buffer
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

if (typeof ArrayBuffer.prototype.resizable !== "undefined") {
    const buf = new ArrayBuffer(8, { maxByteLength: 16 });
    const clone = structuredClone(buf);
    console.log(clone.byteLength);
} else {
    console.log("8");
}
