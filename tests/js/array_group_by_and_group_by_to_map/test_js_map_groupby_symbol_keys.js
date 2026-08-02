// vybe-test: js/array_group_by_and_group_by_to_map/test_js_map_groupby_symbol_keys
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

const s1 = Symbol("group1");
const s2 = Symbol("group2");
const arr = [1, 2, 3, 4];
const grouped = Map.groupBy(arr, x => x % 2 === 0 ? s1 : s2);
__check(__line(grouped.get(s1).join(",") + "|" + grouped.get(s2).join(",")), "2,4|1,3");
