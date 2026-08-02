// vybe-test: js/eval_direct_vs_indirect_scope/test_js_eval_syntax_error_throws_immediately
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

try {
    eval("if (true) {");
} catch (e) {
    __check(__line("Eval SyntaxError: " + (e instanceof SyntaxError)), "Eval SyntaxError: true");
}
