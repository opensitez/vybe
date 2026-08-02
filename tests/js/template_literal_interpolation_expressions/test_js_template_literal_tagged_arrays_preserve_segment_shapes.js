// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_tagged_arrays_preserve_segment_shapes
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

function capture(strings) {
    __check(__line(strings.length), "3");
    __check(__line(strings.raw.length), "3");
    __check(__line(strings[0]), "a\nb");
    __check(__line(strings[1]), "x");
    __check(__line(strings[2]), "y");
    __check(__line(strings.raw[0] === "a\\nb"), "true");
}

capture`a\nb${1}x${2}y`;
