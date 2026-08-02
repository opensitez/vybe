// vybe-test: js/comma_operator_evaluation_and_expression_values/test_js_comma_operator_evaluates_all_operands_left_to_right
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

const log = [];
const res = (log.push(1), log.push(2), log.push(3), "final");
__check(__line(res + "|" + log.join(",")), "final|1,2,3");
