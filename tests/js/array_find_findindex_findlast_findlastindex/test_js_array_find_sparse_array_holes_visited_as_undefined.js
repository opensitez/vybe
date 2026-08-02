// vybe-test: js/array_find_findindex_findlast_findlastindex/test_js_array_find_sparse_array_holes_visited_as_undefined
// origin: languages/js/tests/js/test_js_array_find_findindex_findlast_findlastindex.rs

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
const visited = [];
sparse.find(val => {
    visited.push(String(val));
    return false;
});
__check(__line(visited.join(",")), "1,undefined,3");
