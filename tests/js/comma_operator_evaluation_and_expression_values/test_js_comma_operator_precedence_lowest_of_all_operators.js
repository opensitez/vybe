// vybe-test: js/comma_operator_evaluation_and_expression_values/test_js_comma_operator_precedence_lowest_of_all_operators
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

let a, b;
a = 1, b = 2; // Equivalent to: (a = 1), (b = 2)
__check(__line(`${a}:${b}`), "1:2");
