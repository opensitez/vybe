// vybe-test: js/tagged_template_literal_raw_strings/test_js_tagged_template_raw_property_escapes
// origin: languages/js/tests/js/test_js_tagged_template_literal_raw_strings.rs

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

function rawTag(strings) {
    return strings.raw[0];
}
__check(__line(rawTag`Line1\nLine2\tTabbed`), "Line1\\nLine2\\tTabbed");
