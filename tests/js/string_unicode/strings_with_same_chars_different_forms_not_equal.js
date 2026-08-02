// vybe-test: js/string_unicode/strings_with_same_chars_different_forms_not_equal
// origin: languages/js/tests/js/test_string_unicode.rs

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

const a = "\u00E9";         // composed
const b = "e\u0301";       // decomposed
__check(__line(a === b), "false");
__check(__line(a.normalize("NFC") === b.normalize("NFC")), "true");
