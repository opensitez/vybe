// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_labeled_if_statement_block_break
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

let executed = false;
ifLabel: if (true) {
    executed = true;
    break ifLabel;
    executed = false;
}
__check(__line(executed), "true");
