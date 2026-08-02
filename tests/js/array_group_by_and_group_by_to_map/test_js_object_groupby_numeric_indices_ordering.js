// vybe-test: js/array_group_by_and_group_by_to_map/test_js_object_groupby_numeric_indices_ordering
// origin: languages/js/tests/js/test_js_array_group_by_and_group_by_to_map.rs

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

const arr = [100, 200, 300];
const result = Object.groupBy(arr, (val, idx) => idx);
__check(__line(result[0][0] + "|" + result[1][0]), "100|200");
