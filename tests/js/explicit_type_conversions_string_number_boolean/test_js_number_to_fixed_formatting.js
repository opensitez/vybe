// vybe-test: js/explicit_type_conversions_string_number_boolean/test_js_number_to_fixed_formatting
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
    (123.456).toFixed(2),
    (123.4).toFixed(3),
    (0).toFixed(1)
].join("|")), "123.46|123.400|0.0");
