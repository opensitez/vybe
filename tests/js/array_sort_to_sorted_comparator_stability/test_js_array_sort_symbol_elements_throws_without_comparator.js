// vybe-test: js/array_sort_to_sorted_comparator_stability/test_js_array_sort_symbol_elements_throws_without_comparator
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

const symbols = [Symbol("b"), Symbol("a")];
try {
    symbols.sort(); // String conversion on Symbol throws TypeError!
} catch (e) {
    __check(__line("Sort Symbol Without Comparator TypeError"), "Sort Symbol Without Comparator TypeError");
}
