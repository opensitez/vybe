// vybe-test: js/explicit_type_conversions_string_number_boolean/test_js_parse_int_radix_bounds
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
    parseInt("10", 37); // Radix must be between 2 and 36!
} catch (e) {
    console.log("parseInt Invalid Radix");
}
console.log(isNaN(parseInt("10", 37)));
