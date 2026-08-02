// vybe-test: js/multiline_template_literals_escapes/test_js_multiline_template_literal_newline_preservation
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

const multiline = `Line1
Line2
Line3`;
__check(__line(multiline.split("\n").length), "3");
