// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_labeled_break_in_finally_block_overrides_return
// origin: languages/js/tests/js/test_js_labeled_statements_break_continue_control_flow.rs

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

function fn() {
    lbl: {
        try {
            return "TryReturn";
        } finally {
            break lbl; // Break label in finally overrides return value!
        }
    }
    return "AfterBlock";
}
__check(__line(fn()), "AfterBlock");
