// vybe-test: js/array_hole_length_basics/sparse_array_object_keys_list_only_present_indices
// origin: languages/js/tests/js/test_array_hole_length_basics.rs

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

const arr = [];
arr[1] = "x";
arr[3] = "y";
__check(__line(Object.keys(arr).join(",")), "1,3");
