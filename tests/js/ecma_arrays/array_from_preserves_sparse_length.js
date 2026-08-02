// vybe-test: js/ecma_arrays/array_from_preserves_sparse_length
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

const sparse = [];
sparse[2] = "x";
const arr = Array.from(sparse);
__check(__line(arr.length), "3");
__check(__line(arr[0]), "undefined");
__check(__line(arr[2]), "x");
