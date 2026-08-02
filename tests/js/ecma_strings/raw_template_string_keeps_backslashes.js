// vybe-test: js/ecma_strings/raw_template_string_keeps_backslashes
// origin: languages/js/tests/js/test_ecma_strings.rs

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

const s = String.raw`line1\nline2`;
__check(__line(s === "line1\\nline2"), "true");
__check(__line(s.includes("\\n")), "true");
