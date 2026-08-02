// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_arithmetic_expressions
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

const a = 15, b = 25;
__check(__line(`Sum: ${a + b}, Product: ${a * b}`), "Sum: 40, Product: 375");
