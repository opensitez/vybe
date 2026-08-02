// vybe-test: js/array_map_filter_reduce_reduce_right/test_js_array_filter_sparse_array_holes_skipped
// origin: languages/js/tests/js/test_js_array_map_filter_reduce_reduce_right.rs

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

const sparse = [1, , 3];
const filtered = sparse.filter(() => true);
__check(__line(filtered.length + "|" + filtered.join(",")), "2|1,3");
