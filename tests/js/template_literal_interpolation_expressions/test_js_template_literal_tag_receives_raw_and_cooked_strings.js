// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_tag_receives_raw_and_cooked_strings
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
    __check(__line(strings.raw.length), "2");
    __check(__line(strings[0] === "a\nb"), "true");
    __check(__line(strings.raw[0] === "a\\nb"), "true");
    __check(__line(value), "42");
    __check(__line(strings[1] === "c"), "true");
    __check(__line(strings.raw[1] === "c"), "true");
}
capture`a\nb${42}c`;
