// vybe-test: js/array_hole_length_basics/array_constructor_with_string_single_argument_creates_element
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

const arr = new Array("3");
__check(__line(arr.length), "1");
__check(__line(arr[0]), "3");
