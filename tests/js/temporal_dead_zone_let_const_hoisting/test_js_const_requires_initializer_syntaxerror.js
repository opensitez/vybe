// vybe-test: js/temporal_dead_zone_let_const_hoisting/test_js_const_requires_initializer_syntaxerror
// origin: languages/js/tests/js/test_js_temporal_dead_zone_let_const_hoisting.rs

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
    eval("const uninitialized;");
} catch (e) {
    __check(__line("Const Missing Initializer SyntaxError"), "Const Missing Initializer SyntaxError");
}
