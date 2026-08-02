// vybe-test: js/function_declaration_hoisting_in_blocks/test_js_function_declaration_redeclaring_let_throws_syntaxerror
// origin: languages/js/tests/js/test_js_function_declaration_hoisting_in_blocks.rs

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
    eval("let a = 1; function a() {}");
} catch (e) {
    __check(__line("Redeclare let with Function SyntaxError"), "Redeclare let with Function SyntaxError");
}
