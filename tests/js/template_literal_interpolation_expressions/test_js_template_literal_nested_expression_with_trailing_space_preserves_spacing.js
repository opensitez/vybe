// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_nested_expression_with_trailing_space_preserves_spacing
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

function format(label, value) { return `${label}: ${value}`; }
__check(__line(`${format("a", 1)}|${format("b", 2)}`), "a: 1|b: 2");
