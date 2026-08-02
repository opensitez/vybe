// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_invalid_label_break_throws_syntaxerror
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
    eval("break nonExistentLabel;");
} catch (e) {
    __check(__line("Invalid Break Label SyntaxError"), "Invalid Break Label SyntaxError");
}
