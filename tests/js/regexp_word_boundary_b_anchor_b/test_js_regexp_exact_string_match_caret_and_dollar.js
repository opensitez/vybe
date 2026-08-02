// vybe-test: js/regexp_word_boundary_b_anchor_b/test_js_regexp_exact_string_match_caret_and_dollar
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

const str = "hello";
__check(__line(`${/^hello$/.test("hello")}:${/^hello$/.test("hello world")}`), "true:false");
