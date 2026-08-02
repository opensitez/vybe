// vybe-test: js/array_flat_flatmap_depth_recursion/test_js_array_flat_removes_sparse_array_holes
// origin: languages/js/tests/js/test_js_array_flat_flatmap_depth_recursion.rs

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

const sparse = [1, , 3, [4, , 5]];
const flatSparse = sparse.flat();
__check(__line(flatSparse.length + "|" + flatSparse.join(",")), "4|1,3,4,5");
