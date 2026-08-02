// vybe-test: js/function_length_name_properties_descriptors/test_js_function_name_bound_function_prefix
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

function calc() {}
const boundCalc = calc.bind(null);
const doubleBound = boundCalc.bind(null);
__check(__line(boundCalc.name + "|" + doubleBound.name), "bound calc|bound bound calc");
