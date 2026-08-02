// vybe-test: js/regexp_character_class_escapes_word_digit_space/test_js_regexp_digit_escape_with_unicode_flag
// origin: languages/js/tests/js/test_js_regexp_character_class_escapes_word_digit_space.rs

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

const str = "1₂3"; // ₂ is subscript digit U+2082
__check(__line(str.match(/\d/gu).join(",")), "1,3"); // ASCII digits 1 and 3 match \d!
