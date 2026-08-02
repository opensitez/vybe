// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_typeof_expression_in_interpolation
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

__check(__line(`Type: ${typeof null} / ${typeof [1, 2]} / ${typeof undefined}`), "Type: object / object / undefined");
