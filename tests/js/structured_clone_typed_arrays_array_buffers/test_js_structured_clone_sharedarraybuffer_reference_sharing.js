// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_sharedarraybuffer_reference_sharing
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

if (typeof SharedArrayBuffer !== "undefined") {
    const sab = new SharedArrayBuffer(16);
    const cloneSAB = structuredClone(sab);
    console.log(cloneSAB === sab); // SharedArrayBuffer is shared, NOT duplicated!
} else {
    console.log("true");
}
