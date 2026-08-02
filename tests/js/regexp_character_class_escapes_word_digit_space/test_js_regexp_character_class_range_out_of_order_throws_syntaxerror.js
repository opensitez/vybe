// vybe-test: js/regexp_character_class_escapes_word_digit_space/test_js_regexp_character_class_range_out_of_order_throws_syntaxerror
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

try {
    eval("const re = /[z-a]/;"); // Range out of order (z to a) is a SyntaxError!
} catch (e) {
    __check(__line("Range Out of Order SyntaxError"), "Range Out of Order SyntaxError");
}
