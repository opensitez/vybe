// vybe-test: js/change_array_by_copy_to_reversed_to_spliced_to_sorted_with/test_js_array_toreversed_species_returns_base_array
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

class CustomArray extends Array {}
const ca = new CustomArray(1, 2, 3);
const rev = ca.toReversed();
__check(__line(rev.join(",") + "|isCustom=" + (rev instanceof CustomArray)), "3,2,1|isCustom=false");
