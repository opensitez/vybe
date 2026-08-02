// vybe-test: js/explicit_type_conversions_string_number_boolean/test_js_number_constructor_primitives
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
    Number("42"),
    Number("  3.14  "),
    Number(""),
    Number("   "),
    Number(true),
    Number(false),
    Number(null),
    Number(undefined),
    Number("invalid")
].join("|")), "42|3.14|0|0|1|0|0|NaN|NaN");
