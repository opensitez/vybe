// vybe-test: js/array_hole_length_basics/shift_on_sparse_array_moves_present_elements_left
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
const first = arr.shift();
__check(__line(first === undefined), "true");
__check(__line(arr.length), "3");
__check(__line(Object.keys(arr).join(",")), "0,2");
