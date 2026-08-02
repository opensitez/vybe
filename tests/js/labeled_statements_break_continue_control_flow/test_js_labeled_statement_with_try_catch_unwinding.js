// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_labeled_statement_with_try_catch_unwinding
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

const log = [];
lbl: {
    try {
        log.push("Try");
        throw new Error("Err");
    } catch (e) {
        log.push("Catch");
        break lbl;
    } finally {
        log.push("Finally");
    }
    log.push("Unreachable");
}
__check(__line(log.join(",")), "Try,Catch,Finally");
