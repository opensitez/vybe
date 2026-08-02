// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_expression_evaluation_order_with_comma_operator
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

let x = 0;
__check(__line(`value=${(x++, x += 10)}`), "value=11");
__check(__line(x), "11");
