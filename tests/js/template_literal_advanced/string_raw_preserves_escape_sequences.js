// vybe-test: js/template_literal_advanced/string_raw_preserves_escape_sequences
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

const path = String.raw`C:\Users\name\file.txt`;
__check(__line(path), "C:\\Users\\name\\file.txt");
