// vybe-test: js/array_holes_sparse/fill_converts_hole_to_value
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

const arr = [1, , 3];
arr.fill(0, 1, 2);
__check(__line(arr.length), "3");
__check(__line(1 in arr), "true");
__check(__line(arr.join(",")), "1,0,3");
