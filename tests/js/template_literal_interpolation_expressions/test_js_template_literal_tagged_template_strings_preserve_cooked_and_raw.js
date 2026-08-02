// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_tagged_template_strings_preserve_cooked_and_raw
// origin: languages/js/tests/js/test_js_template_literal_interpolation_expressions.rs

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

function capture(strings, value) {
    __check(__line(strings.length), "2");
    __check(__line(strings[0] === "a\nb"), "true");
    __check(__line(strings.raw[0] === "a\\nb"), "true");
    return value;
}
__check(__line(capture`a\nb${41 + 1}c`), "42");
