// vybe-test: js/control_flow_advanced/test_js_control_flow_switch_body_shared_lexical_scope_redeclaration_error
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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
    eval("switch (1) { case 1: let x = 10; break; case 2: let x = 20; break; }");
} catch (e) {
    __check(__line("Switch Single Lexical Scope SyntaxError"), "Switch Single Lexical Scope SyntaxError");
}
