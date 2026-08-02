// vybe-test: js/array_group_by_and_group_by_to_map/test_js_map_groupby_nan_keys_grouped_together
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

const arr = [1, 2, 3];
const grouped = Map.groupBy(arr, () => NaN);
__check(__line(grouped.size + "|" + grouped.get(NaN).length), "1|3"); // NaN key uses SameValueZero equality!
