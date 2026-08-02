// vybe-test: js/temporal_dead_zone_let_const_hoisting/test_js_redeclaring_var_as_let_in_same_scope_throws_syntaxerror
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
    eval("var b = 1; let b = 2;");
} catch (e) {
    __check(__line("Redeclare var as let SyntaxError"), "Redeclare var as let SyntaxError");
}
