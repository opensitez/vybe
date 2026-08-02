// vybe-test: js/array_flat_flatmap_depth_recursion/test_js_array_flatmap_non_callable_callback_throws
// origin: languages/js/tests/js/test_js_array_flat_flatmap_depth_recursion.rs

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
    [1, 2].flatMap("not_callable");
} catch (e) {
    __check(__line("flatMap Non-Callable TypeError"), "flatMap Non-Callable TypeError");
}
