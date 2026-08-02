// vybe-test: js/map_groupby_object_groupby_utilities/test_js_map_groupby_nan_key_grouping
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

const values = [NaN, 10, NaN, 20];
const grouped = Map.groupBy(values, val => Number.isNaN(val) ? NaN : "number");
__check(__line(grouped.get(NaN).length + "|" + grouped.get("number").join(",")), "2|10,20");
