// vybe-test: js/multiline_template_literals_escapes/test_js_template_literal_unicode_escape_sequences
// origin: languages/js/tests/js/test_js_multiline_template_literals_escapes.rs

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

const unicode = `\u0041\u0042\u0043`;
__check(__line(unicode), "ABC");
