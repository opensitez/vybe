// vybe-test: js/array_hole_length_basics/sparse_array_flat_removes_holes
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

const arr = [1, , 3, , 5];
__check(__line(arr.flat().length + "|" + arr.flat().join(",")), "3|1,3,5");
