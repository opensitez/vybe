// vybe-test: js/typed_array_subarray_slice_set_copywithin/test_js_typedarray_set_coercion_of_input_elements
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

const arr = new Uint8Array(2);
arr.set(["10", 256]); // "10" -> 10, 256 -> 0
__check(__line(arr.join(",")), "10,0");
