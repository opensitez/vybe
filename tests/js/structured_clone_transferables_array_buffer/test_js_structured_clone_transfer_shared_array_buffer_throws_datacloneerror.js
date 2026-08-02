// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_shared_array_buffer_throws_datacloneerror
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

if (typeof SharedArrayBuffer !== "undefined") {
    const sab = new SharedArrayBuffer(16);
    try {
        structuredClone(sab, { transfer: [sab] }); // SharedArrayBuffer cannot be transferred!
    } catch (e) {
        console.log("DataCloneError Transfer SharedArrayBuffer");
    }
} else {
    console.log("DataCloneError Transfer SharedArrayBuffer");
}
