// vybe-test: js/array_group_by_and_group_by_to_map/test_js_map_groupby_non_callable_callback_throws_typeerror
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

try {
    Map.groupBy([1, 2], null);
} catch (e) {
    __check(__line("Map.groupBy Callback TypeError"), "Map.groupBy Callback TypeError");
}
