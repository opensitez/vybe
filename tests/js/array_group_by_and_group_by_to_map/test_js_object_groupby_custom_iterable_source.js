// vybe-test: js/array_group_by_and_group_by_to_map/test_js_object_groupby_custom_iterable_source
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

const set = new Set(["apple", "banana", "avocado"]);
const grouped = Object.groupBy(set, word => word[0]);
__check(__line(grouped.a.join(",") + "|" + grouped.b.join(",")), "apple,avocado|banana");
