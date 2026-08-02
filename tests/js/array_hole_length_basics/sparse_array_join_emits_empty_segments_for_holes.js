// vybe-test: js/array_hole_length_basics/sparse_array_join_emits_empty_segments_for_holes
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
__check(__line(arr.join(",")), ",x,,y");
