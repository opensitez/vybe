// vybe-test: js/function_length_name_properties_descriptors/test_js_function_name_symbol_computed_properties
// origin: languages/js/tests/js/test_js_function_length_name_properties_descriptors.rs

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

const sym = Symbol("mySym");
const obj = {
    [sym]() {}
};
__check(__line(obj[sym].name), "[mySym]");
