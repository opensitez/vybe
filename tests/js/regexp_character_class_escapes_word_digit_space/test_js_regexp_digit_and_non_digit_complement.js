// vybe-test: js/regexp_character_class_escapes_word_digit_space/test_js_regexp_digit_and_non_digit_complement
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

const str = "123abc";
__check(__line(str.replace(/\d/g, "#") + "|" + str.replace(/\D/g, "*")), "###abc|123***");
