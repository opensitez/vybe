// vybe-test: js/eval_direct_vs_indirect_scope/test_js_eval_non_string_argument_returns_argument_as_is
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

__check(__line(eval(123) + "|" + eval(true) + "|" + (eval(null) === null)), "123|true|true");
