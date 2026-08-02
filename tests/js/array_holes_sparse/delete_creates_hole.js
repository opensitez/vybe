// vybe-test: js/array_holes_sparse/delete_creates_hole
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

const arr = [1, 2, 3];
delete arr[1];
__check(__line(arr.length), "3");  // unchanged
__check(__line(1 in arr), "false");    // false — hole
__check(__line(arr[1]), "undefined");      // undefined
