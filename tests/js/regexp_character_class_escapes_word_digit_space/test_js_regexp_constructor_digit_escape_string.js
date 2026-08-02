// vybe-test: js/regexp_character_class_escapes_word_digit_space/test_js_regexp_constructor_digit_escape_string
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

const re = new RegExp("\\d+");
__check(__line(re.test("123")), "true");
