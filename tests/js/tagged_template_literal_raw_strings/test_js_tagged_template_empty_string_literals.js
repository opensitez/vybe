// vybe-test: js/tagged_template_literal_raw_strings/test_js_tagged_template_empty_string_literals
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

function inspect(strings, ...values) {
    return strings.length + "|" + values.length;
}
__check(__line(inspect``), "1|0");
