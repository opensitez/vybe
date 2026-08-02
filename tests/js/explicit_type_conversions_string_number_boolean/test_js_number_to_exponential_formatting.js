// vybe-test: js/explicit_type_conversions_string_number_boolean/test_js_number_to_exponential_formatting
// origin: languages/js/tests/js/test_js_explicit_type_conversions_string_number_boolean.rs

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

__check(__line([
    (123456).toExponential(2),
    (0.005).toExponential(1)
].join("|")), "1.23e+5|5.0e-3");
