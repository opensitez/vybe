// vybe-test: js/eval_direct_vs_indirect_scope/test_js_eval_string_wrapper_object_returns_as_is
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

const strObj = new String("2 + 2");
const res = eval(strObj);
__check(__line(typeof res + "|" + (res === strObj)), "object|true");
