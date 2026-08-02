// vybe-test: js/comma_operator_evaluation_and_expression_values/test_js_comma_operator_exception_aborts_subsequent_operands
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

let step = 0;
try {
    const res = (step = 1, (() => { throw new Error("abort"); })(), step = 2);
} catch (e) {
    __check(__line(e.message + "|step=" + step), "abort|step=1");
}
