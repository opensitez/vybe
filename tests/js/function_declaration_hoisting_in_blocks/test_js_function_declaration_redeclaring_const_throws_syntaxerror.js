// vybe-test: js/function_declaration_hoisting_in_blocks/test_js_function_declaration_redeclaring_const_throws_syntaxerror
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
    eval("const c = 1; function c() {}");
} catch (e) {
    __check(__line("Redeclare const with Function SyntaxError"), "Redeclare const with Function SyntaxError");
}
