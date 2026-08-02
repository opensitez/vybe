// vybe-test: js/try_catch_finally_return_override_control_flow/test_js_finally_return_expression_evaluated_before_cleanup
// origin: languages/js/tests/js/test_js_try_catch_finally_return_override_control_flow.rs

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

let sideEffect = 0;
function getVal() {
    sideEffect++;
    return sideEffect;
}
function fn() {
    try {
        return 99;
    } finally {
        return getVal();
    }
}
__check(__line(fn() + "|SideEffect=" + sideEffect), "1|SideEffect=1");
