// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_labeled_statement_scope_shadowing_same_name_different_blocks
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

const res = [];
blockA: {
    res.push("A1");
}
blockA: {
    res.push("A2");
}
__check(__line(res.join(",")), "A1,A2");
