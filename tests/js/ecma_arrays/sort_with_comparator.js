// vybe-test: js/ecma_arrays/sort_with_comparator
// origin: languages/js/tests/js/test_ecma_arrays.rs

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

const arr = [3, 1, 4, 1, 5];
arr.sort((a, b) => b - a);
__check(__line(arr.join(",")), "5,4,3,1,1");
