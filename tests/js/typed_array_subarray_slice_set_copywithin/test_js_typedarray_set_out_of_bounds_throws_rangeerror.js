// vybe-test: js/typed_array_subarray_slice_set_copywithin/test_js_typedarray_set_out_of_bounds_throws_rangeerror
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

const dest = new Uint8Array(3);
try {
    dest.set([1, 2, 3], 2); // Exceeds length 3!
} catch (e) {
    __check(__line("TypedArray Set RangeError"), "TypedArray Set RangeError");
}
