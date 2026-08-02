// vybe-test: js/comma_operator_evaluation_and_expression_values/test_js_comma_operator_with_throw_expression_syntax_error
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

try {
    eval("const x = (1, throw new Error());"); // throw statement inside expression is a SyntaxError!
} catch (e) {
    __check(__line("Comma Operator Throw SyntaxError"), "Comma Operator Throw SyntaxError");
}
