// vybe-test: js/global_builtin_edges/eval_can_define_local_binding_in_current_scope
// origin: languages/js/tests/js/test_global_builtin_edges.rs

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

eval("var fromEval = 42;");
__check(__line(fromEval), "42");
