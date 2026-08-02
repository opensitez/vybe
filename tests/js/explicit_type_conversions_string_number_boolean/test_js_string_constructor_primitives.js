// vybe-test: js/explicit_type_conversions_string_number_boolean/test_js_string_constructor_primitives
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
    String(123),
    String(true),
    String(false),
    String(null),
    String(undefined),
    String(10n),
    String(Symbol("id"))
].join("|")), "123|true|false|null|undefined|10|Symbol(id)");
