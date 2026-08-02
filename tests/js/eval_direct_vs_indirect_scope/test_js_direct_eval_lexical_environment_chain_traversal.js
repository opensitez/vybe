// vybe-test: js/eval_direct_vs_indirect_scope/test_js_direct_eval_lexical_environment_chain_traversal
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

function outer() {
    const a = 10;
    function inner() {
        const b = 20;
        return eval("a + b");
    }
    return inner();
}
__check(__line(outer()), "30");
