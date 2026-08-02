// vybe-test: js/map_groupby_object_groupby_utilities/test_js_object_groupby_sparse_array_holes_visited
// origin: languages/js/tests/js/test_js_map_groupby_object_groupby_utilities.rs

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
const grouped = Object.groupBy(sparse, item => item === undefined ? "undef" : "def");
__check(__line(grouped.def.join(",") + "|countUndef=" + grouped.undef.length), "1,3|countUndef=1");
