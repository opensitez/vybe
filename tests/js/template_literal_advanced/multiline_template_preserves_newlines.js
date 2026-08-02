// vybe-test: js/template_literal_advanced/multiline_template_preserves_newlines
// origin: languages/js/tests/js/test_template_literal_advanced.rs

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

const text = `line1
line2
line3`;
const lines = text.split("\n");
__check(__line(lines.length), "3");
__check(__line(lines[1]), "line2");
