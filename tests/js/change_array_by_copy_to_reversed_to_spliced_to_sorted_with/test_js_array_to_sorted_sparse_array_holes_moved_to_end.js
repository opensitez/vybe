// vybe-test: js/change_array_by_copy_to_reversed_to_spliced_to_sorted_with/test_js_array_to_sorted_sparse_array_holes_moved_to_end
// origin: languages/js/tests/js/test_js_change_array_by_copy_to_reversed_to_spliced_to_sorted_with.rs

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

const sparse = [2, , 1];
const sorted = sparse.toSorted();
__check(__line(sorted.join(",") + "|len=" + sorted.length), "1,2,|len=3");
