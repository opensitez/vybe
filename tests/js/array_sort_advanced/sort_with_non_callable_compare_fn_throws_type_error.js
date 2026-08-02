// vybe-test: js/array_sort_advanced/sort_with_non_callable_compare_fn_throws_type_error
// origin: languages/js/tests/js/test_array_sort_advanced.rs

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
    [3, 1].sort("not a function");
} catch (e) {
    __check(__line(e.name), "TypeError");
}
