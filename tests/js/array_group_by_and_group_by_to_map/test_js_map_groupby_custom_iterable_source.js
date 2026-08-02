// vybe-test: js/array_group_by_and_group_by_to_map/test_js_map_groupby_custom_iterable_source
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

const generator = function*() { yield 10; yield 20; yield 15; };
const grouped = Map.groupBy(generator(), x => x >= 15);
__check(__line(grouped.get(true).join(",") + "|" + grouped.get(false).join(",")), "20,15|10");
