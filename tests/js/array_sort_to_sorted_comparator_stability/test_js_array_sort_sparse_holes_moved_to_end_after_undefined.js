// vybe-test: js/array_sort_to_sorted_comparator_stability/test_js_array_sort_sparse_holes_moved_to_end_after_undefined
// origin: languages/js/tests/js/test_js_array_sort_to_sorted_comparator_stability.rs

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

const sparse = [2, , 1, undefined, 3];
sparse.sort((a, b) => a - b);
__check(__line(sparse.length + "|" + sparse.map(x => String(x)).join(",")), "5|1,2,3,undefined,undefined");
