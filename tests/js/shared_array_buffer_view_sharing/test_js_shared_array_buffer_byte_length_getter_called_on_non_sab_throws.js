// vybe-test: js/shared_array_buffer_view_sharing/test_js_shared_array_buffer_byte_length_getter_called_on_non_sab_throws
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

const getter = Object.getOwnPropertyDescriptor(SharedArrayBuffer.prototype, "byteLength").get;
try {
    getter.call(new ArrayBuffer(8));
} catch (e) {
    __check(__line("byteLength Non-SAB TypeError"), "byteLength Non-SAB TypeError");
}
