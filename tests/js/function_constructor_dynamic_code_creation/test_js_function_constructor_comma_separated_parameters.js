// vybe-test: js/function_constructor_dynamic_code_creation/test_js_function_constructor_comma_separated_parameters
// origin: languages/js/tests/js/test_js_function_constructor_dynamic_code_creation.rs

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

const sum3 = new Function("a, b, c", "return a + b + c;");
__check(__line(sum3(1, 2, 3)), "6");
