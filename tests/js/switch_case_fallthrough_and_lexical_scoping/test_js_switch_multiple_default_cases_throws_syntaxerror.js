// vybe-test: js/switch_case_fallthrough_and_lexical_scoping/test_js_switch_multiple_default_cases_throws_syntaxerror
// origin: languages/js/tests/js/test_js_switch_case_fallthrough_and_lexical_scoping.rs

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
    eval("switch(1) { default: break; default: break; }");
} catch (e) {
    __check(__line("Multiple Defaults SyntaxError"), "Multiple Defaults SyntaxError");
}
