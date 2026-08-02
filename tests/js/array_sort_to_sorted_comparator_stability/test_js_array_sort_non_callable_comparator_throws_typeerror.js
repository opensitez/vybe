// vybe-test: js/array_sort_to_sorted_comparator_stability/test_js_array_sort_non_callable_comparator_throws_typeerror
// origin: languages/js/tests/js/test_js_array_sort_to_sorted_comparator_stability.rs

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
    [1, 2].sort(123);
} catch (e) {
    __check(__line("Sort Invalid Comparator TypeError"), "Sort Invalid Comparator TypeError");
}
