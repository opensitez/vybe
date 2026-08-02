// vybe-test: js/array_hole_length_basics/array_length_extension_creates_new_holes
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

const arr = [1, 2];
arr.length = 4;
__check(__line(arr.length), "4");
__check(__line(2 in arr), "false");
__check(__line(arr[3] === undefined), "true");
