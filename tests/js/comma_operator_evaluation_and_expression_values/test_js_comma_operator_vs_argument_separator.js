// vybe-test: js/comma_operator_evaluation_and_expression_values/test_js_comma_operator_vs_argument_separator
// origin: languages/js/tests/js/test_js_comma_operator_evaluation_and_expression_values.rs

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

function fn(x) { return x; }
__check(__line(fn((10, 20))), "20"); // Parentheses enforce comma operator, passing 20 to single parameter x!
