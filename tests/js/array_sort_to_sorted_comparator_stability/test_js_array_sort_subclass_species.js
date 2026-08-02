// vybe-test: js/array_sort_to_sorted_comparator_stability/test_js_array_sort_subclass_species
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

class CustomArray extends Array {}
const ca = new CustomArray(3, 1, 2);
ca.sort();
__check(__line(ca.join(",") + "|isCustom=" + (ca instanceof CustomArray)), "1,2,3|isCustom=true");
