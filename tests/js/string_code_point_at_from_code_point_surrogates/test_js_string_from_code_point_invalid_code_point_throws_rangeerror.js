// vybe-test: js/string_code_point_at_from_code_point_surrogates/test_js_string_from_code_point_invalid_code_point_throws_rangeerror
// origin: languages/js/tests/js/test_js_string_code_point_at_from_code_point_surrogates.rs

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
    String.fromCodePoint(0x110000); // Beyond U+10FFFF max code point!
} catch (e) {
    __check(__line("fromCodePoint RangeError"), "fromCodePoint RangeError");
}
