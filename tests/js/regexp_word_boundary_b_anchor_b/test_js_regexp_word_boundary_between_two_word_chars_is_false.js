// vybe-test: js/regexp_word_boundary_b_anchor_b/test_js_regexp_word_boundary_between_two_word_chars_is_false
// origin: languages/js/tests/js/test_js_regexp_word_boundary_b_anchor_b.rs

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

__check(__line(/\b/.test("aa")), "true");
