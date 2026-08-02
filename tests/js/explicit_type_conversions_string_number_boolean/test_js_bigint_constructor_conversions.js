// vybe-test: js/explicit_type_conversions_string_number_boolean/test_js_bigint_constructor_conversions
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
    BigInt(100).toString(),
    BigInt("200").toString(),
    BigInt("0b1010").toString(),
    BigInt("0xff").toString(),
    BigInt(true).toString(),
    BigInt(false).toString()
].join("|")), "100|200|10|255|1|0");
