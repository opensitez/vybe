// vybe-test: js/array_map_filter_reduce_reduce_right/test_js_array_map_non_callable_callback_throws
// origin: languages/js/tests/js/test_js_array_map_filter_reduce_reduce_right.rs

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
    [1, 2].map("not_fn");
} catch (e) {
    __check(__line("Map Non-Callable TypeError"), "Map Non-Callable TypeError");
}
