// vybe-test: js/array_holes_sparse/at_and_inclusive_checks
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
__check(__line(arr.at(1)), "undefined");
__check(__line(arr[1]), "undefined");
__check(__line(1 in arr), "false");
