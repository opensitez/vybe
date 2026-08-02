// vybe-test: js/typed_array_subarray_slice_set_copywithin/test_js_typedarray_slice_returns_new_buffer
// origin: languages/js/tests/js/test_js_typed_array_subarray_slice_set_copywithin.rs

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

const u8 = new Uint8Array([1, 2]);
const sliced = u8.slice();
__check(__line(sliced.buffer !== u8.buffer), "true");
