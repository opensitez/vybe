// vybe-test: js/array_sort_to_sorted_comparator_stability/test_js_array_sort_symbol_elements_with_custom_comparator
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

const s1 = Symbol("a"), s2 = Symbol("b");
const symbols = [s2, s1];
symbols.sort((a, b) => a.description.localeCompare(b.description));
__check(__line(symbols.map(s => s.description).join(",")), "a,b");
