// vybe-test: js/explicit_type_conversions_string_number_boolean/test_js_parse_float_conversions
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
    parseFloat("3.14"),
    parseFloat("314e-2"),
    parseFloat("10.5.6"),
    parseFloat("  42  "),
    parseFloat("text")
].join("|")), "3.14|3.14|10.5|42|NaN");
