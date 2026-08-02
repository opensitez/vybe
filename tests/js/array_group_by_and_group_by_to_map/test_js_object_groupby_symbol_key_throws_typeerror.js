// vybe-test: js/array_group_by_and_group_by_to_map/test_js_object_groupby_symbol_key_throws_typeerror
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
    Object.groupBy([1], () => Symbol("sym")); // Object.groupBy coerces key to string, Symbol throws TypeError!
} catch (e) {
    __check(__line("Object.groupBy Symbol Key TypeError"), "Object.groupBy Symbol Key TypeError");
}
