// vybe-test: js/explicit_type_conversions_string_number_boolean/test_js_parse_int_radix_conversions
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
    parseInt("42", 10),
    parseInt("1010", 2),
    parseInt("ff", 16),
    parseInt("077", 8),
    parseInt("100px", 10),
    parseInt("abc", 10)
].join("|")), "42|10|255|77|100|NaN");
