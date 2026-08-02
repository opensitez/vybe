// vybe-test: js/explicit_type_conversions_string_number_boolean/test_js_bigint_constructor_symbol_throws_typeerror
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

try {
    BigInt(Symbol("foo"));
} catch (e) {
    __check(__line(e.name), "TypeError");
}
