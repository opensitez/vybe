// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_multiple_labels_on_single_statement
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

let hit = false;
label1: label2: {
    hit = true;
    break label1;
}
__check(__line(hit), "true");
