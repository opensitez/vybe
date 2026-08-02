// vybe-test: js/typed_array_subarray_slice_set_copywithin/test_js_typedarray_set_copy_from_array
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

const dest = new Uint8Array(5);
dest.set([10, 20, 30], 1); // Set starting at index 1
__check(__line(dest.join(",")), "0,10,20,30,0");
