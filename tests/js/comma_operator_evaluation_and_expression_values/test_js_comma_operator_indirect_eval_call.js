// vybe-test: js/comma_operator_evaluation_and_expression_values/test_js_comma_operator_indirect_eval_call
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

const x = "globalScope";
(() => {
    const x = "localScope";
    const indirectEval = (0, eval); // (0, eval) performs indirect eval in global scope!
    __check(__line(indirectEval("x")), "globalScope");
})();
