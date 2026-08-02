// vybe-test: js/array_map_filter_reduce_reduce_right/test_js_array_filter_subclass_species
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

class CustomArray extends Array {}
const ca = new CustomArray(1, 2, 3, 4);
const filtered = ca.filter(x => x > 2);
__check(__line(filtered.join(",") + "|isCustom=" + (filtered instanceof CustomArray)), "3,4|isCustom=true");
