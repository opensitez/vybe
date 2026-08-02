// vybe-test: js/array_hole_length_basics/sparse_array_reduce_uses_first_present_element_without_initial
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
arr[2] = 5;
arr[4] = 7;
__check(__line(arr.reduce((acc, value) => acc + value)), "12");
