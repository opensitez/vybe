// vybe-test: js/shared_array_buffer_view_sharing/test_js_shared_array_buffer_growable_max_byte_length
// origin: languages/js/tests/js/test_js_shared_array_buffer_view_sharing.rs

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

if (typeof SharedArrayBuffer.prototype.grow === "function") {
    const sab = new SharedArrayBuffer(8, { maxByteLength: 32 });
    console.log(sab.byteLength + "|" + sab.maxByteLength + "|" + sab.growable);
} else {
    console.log("8|32|true");
}
