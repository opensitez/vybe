// vybe-test: js/array_hole_length_basics/pop_on_trailing_hole_returns_undefined_and_shrinks_length
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

const arr = [1, 2, 3];
arr.length = 4;
__check(__line(arr.pop() === undefined), "true");
__check(__line(arr.length), "3");
__check(__line(arr.join(",")), "1,2,3");
