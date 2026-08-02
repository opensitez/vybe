// vybe-test: js/array_splice_to_spliced_slice_mutation/test_js_array_slice_shallow_copy_preserves_object_references
// origin: languages/js/tests/js/test_js_array_splice_to_spliced_slice_mutation.rs

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

const obj = { id: 1 };
const arr = [obj];
const sliced = arr.slice();
__check(__line(sliced[0] === obj), "true");
