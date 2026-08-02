// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_duplicate_label_in_nested_scope_throws_syntaxerror
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

try {
    eval("label: { label: { break label; } }");
} catch (e) {
    __check(__line("Duplicate Label SyntaxError"), "Duplicate Label SyntaxError");
}
