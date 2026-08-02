// vybe-test: js/array_sort_to_sorted_comparator_stability/test_js_array_sort_frozen_array_throws_in_strict
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

const frozen = Object.freeze([3, 1, 2]);
try {
    "use strict";
    frozen.sort();
} catch (e) {
    __check(__line("Sort Frozen Array TypeError"), "Sort Frozen Array TypeError");
}
