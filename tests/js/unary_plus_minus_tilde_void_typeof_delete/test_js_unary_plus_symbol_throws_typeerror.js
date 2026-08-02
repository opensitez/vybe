// vybe-test: js/unary_plus_minus_tilde_void_typeof_delete/test_js_unary_plus_symbol_throws_typeerror
// origin: languages/js/tests/js/test_js_unary_plus_minus_tilde_void_typeof_delete.rs

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
    console.log(+Symbol("x"));
} catch (e) {
    console.log("Unary Plus Symbol TypeError");
}
