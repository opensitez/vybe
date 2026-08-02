// vybe-test: js/array_buffer_slice_transfer_resizable/test_js_arraybuffer_constructor_negative_length_throws_rangeerror
// origin: languages/js/tests/js/test_js_array_buffer_slice_transfer_resizable.rs

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

try {
    new ArrayBuffer(-10);
} catch (e) {
    __check(__line("Negative Buffer Length RangeError"), "Negative Buffer Length RangeError");
}
