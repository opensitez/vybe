// vybe-test: js/regexp_character_class_escapes_word_digit_space/test_js_regexp_whitespace_escape_non_breaking_space
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

const str = "A\u00A0B"; // Non-breaking space \u00A0
__check(__line(str.match(/\s/g).length), "1");
