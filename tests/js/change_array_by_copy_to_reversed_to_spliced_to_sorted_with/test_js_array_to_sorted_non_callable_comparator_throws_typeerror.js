// vybe-test: js/change_array_by_copy_to_reversed_to_spliced_to_sorted_with/test_js_array_to_sorted_non_callable_comparator_throws_typeerror
// origin: languages/js/tests/js/test_js_change_array_by_copy_to_reversed_to_spliced_to_sorted_with.rs

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
    [1, 2].toSorted("not_a_function");
} catch (e) {
    __check(__line("toSorted Comparator TypeError"), "toSorted Comparator TypeError");
}
