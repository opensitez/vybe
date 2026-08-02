// vybe-test: js/regex_string_methods/unicode_property_escape_basic
// origin: languages/js/tests/js/test_regex_string_methods.rs

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

const re = /\p{Lu}/u; // Uppercase letter
__check(__line(re.test("A")), "true");
__check(__line(re.test("a")), "false");
