// vybe-test: js/function_length_name_properties_descriptors/test_js_function_length_counts_parameters_before_default
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

function fn1(a, b, c) {}
function fn2(a, b = 1, c) {}
function fn3(a = 1, b, c) {}
__check(__line(`${fn1.length}:${fn2.length}:${fn3.length}`), "3:1:0");
