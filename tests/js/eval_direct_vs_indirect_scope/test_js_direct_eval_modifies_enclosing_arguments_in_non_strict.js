// vybe-test: js/eval_direct_vs_indirect_scope/test_js_direct_eval_modifies_enclosing_arguments_in_non_strict
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

function fn(a) {
    eval("a = 10;");
    return a;
}
__check(__line(fn(5)), "10");
