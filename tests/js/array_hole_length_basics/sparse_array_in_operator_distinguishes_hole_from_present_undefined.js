// vybe-test: js/array_hole_length_basics/sparse_array_in_operator_distinguishes_hole_from_present_undefined
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

const arr = [undefined, , undefined];
__check(__line(0 in arr), "true");
__check(__line(1 in arr), "false");
__check(__line(2 in arr), "true");
