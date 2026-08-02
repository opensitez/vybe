// vybe-test: js/array_flat_flatmap_depth_recursion/test_js_array_flat_subclass_species
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

class CustomArray extends Array {}
const ca = new CustomArray(1, [2, 3]);
const flat = ca.flat();
__check(__line(flat.join(",") + "|isCustom=" + (flat instanceof CustomArray)), "1,2,3|isCustom=true");
