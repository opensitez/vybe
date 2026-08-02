// vybe-test: js/array_hole_length_basics/array_constructor_with_multiple_arguments_creates_elements
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

const arr = new Array(1, 2, 3);
__check(__line(arr.length), "3");
__check(__line(arr.join(",")), "1,2,3");
