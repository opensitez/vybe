// vybe-test: js/array_holes_sparse/hole_in_array_literal
// origin: languages/js/tests/js/test_array_holes_sparse.rs

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

const arr = [1, , 3]; // hole at index 1
__check(__line(arr.length), "3");
__check(__line(arr[1]), "undefined");       // undefined
__check(__line(1 in arr), "false");     // false — hole, not undefined
