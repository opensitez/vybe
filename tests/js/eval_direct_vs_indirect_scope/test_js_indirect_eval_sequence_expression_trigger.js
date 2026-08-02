// vybe-test: js/eval_direct_vs_indirect_scope/test_js_indirect_eval_sequence_expression_trigger
// origin: languages/js/tests/js/test_js_eval_direct_vs_indirect_scope.rs

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

var x = 1;
function fn() {
    var x = 2;
    return (0, eval)("x"); // (0, eval) is an indirect call -> evaluates in global scope!
}
__check(__line(fn()), "1");
