// vybe-test: js/comma_operator_evaluation_and_expression_values/test_js_comma_operator_in_variable_declarations_without_parentheses
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

let a = 1, b = 2, c = 3; // Declares three variables, NOT comma operator expression!
__check(__line(`${a}:${b}:${c}`), "1:2:3");
