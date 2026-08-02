// vybe-test: js/map_groupby_object_groupby_utilities/test_js_object_groupby_non_callable_callback_throws
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

try {
    Object.groupBy([1, 2], "not_a_function");
} catch (e) {
    __check(__line("Object.groupBy Non-Callable TypeError"), "Object.groupBy Non-Callable TypeError");
}
