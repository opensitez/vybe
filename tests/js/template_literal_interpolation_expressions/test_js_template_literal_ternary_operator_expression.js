// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_ternary_operator_expression
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

const isMember = true;
__check(__line(`Fee: $${isMember ? 10 : 50}`), "Fee: $10");
