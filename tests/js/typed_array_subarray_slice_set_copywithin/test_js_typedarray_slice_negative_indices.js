// vybe-test: js/typed_array_subarray_slice_set_copywithin/test_js_typedarray_slice_negative_indices
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

const arr = new Int32Array([10, 20, 30, 40]);
const sliced = arr.slice(-2);
__check(__line(sliced.join(",")), "30,40");
