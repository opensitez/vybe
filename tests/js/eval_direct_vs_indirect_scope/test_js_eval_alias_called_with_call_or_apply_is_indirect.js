// vybe-test: js/eval_direct_vs_indirect_scope/test_js_eval_alias_called_with_call_or_apply_is_indirect
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

var g = "GlobalVal";
function test() {
    var g = "LocalVal";
    return eval.call(null, "g");
}
__check(__line(test()), "GlobalVal");
