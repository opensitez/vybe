// vybe-test: js/explicit_type_conversions_string_number_boolean/test_js_boolean_constructor_falsy_values
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
    Boolean(false),
    Boolean(0),
    Boolean(-0),
    Boolean(0n),
    Boolean(""),
    Boolean(null),
    Boolean(undefined),
    Boolean(NaN)
].every(val => val === false)), "true");
