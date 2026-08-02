// vybe-test: js/regexp_unicode_sets_v_flag/test_js_regexp_unicode_sets_negated_class_with_strings_throws
// origin: languages/js/tests/js/test_js_regexp_unicode_sets_v_flag.rs

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
    eval("const re = /[^\\p{Basic_Emoji}]/v;");
} catch (e) {
    __check(__line("Negated String Property Class SyntaxError"), "Negated String Property Class SyntaxError");
}
