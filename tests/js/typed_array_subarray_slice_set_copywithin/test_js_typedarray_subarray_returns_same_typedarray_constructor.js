// vybe-test: js/typed_array_subarray_slice_set_copywithin/test_js_typedarray_subarray_returns_same_typedarray_constructor
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

const i32 = new Int32Array([1, 2, 3]);
const sub = i32.subarray(1);
__check(__line(sub instanceof Int32Array), "true");
