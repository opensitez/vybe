// vybe-test: js/comma_operator_evaluation_and_expression_values/test_js_comma_operator_in_while_loop_condition
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

let i = 0;
const log = [];
while (log.push(i), i < 2) {
    i++;
}
console.log(log.join(","));
