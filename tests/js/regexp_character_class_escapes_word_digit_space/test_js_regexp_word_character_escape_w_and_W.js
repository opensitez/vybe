// vybe-test: js/regexp_character_class_escapes_word_digit_space/test_js_regexp_word_character_escape_w_and_W
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

const str = "A_1 !@#";
__check(__line(str.match(/\w/g).join(",") + "|" + str.match(/\W/g).join(",")), "A,_,1| ,!,@,#");
