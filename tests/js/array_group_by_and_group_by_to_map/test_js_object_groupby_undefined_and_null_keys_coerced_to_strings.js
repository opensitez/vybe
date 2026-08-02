// vybe-test: js/array_group_by_and_group_by_to_map/test_js_object_groupby_undefined_and_null_keys_coerced_to_strings
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

const arr = [1, 2];
const grouped = Object.groupBy(arr, x => x === 1 ? null : undefined);
__check(__line(grouped["null"].length + "|" + grouped["undefined"].length), "1|1");
