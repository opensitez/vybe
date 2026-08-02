// vybe-test: js/temporal_dead_zone_let_const_hoisting/test_js_tdz_access_before_const_declaration_throws_referenceerror
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
    eval("console.log(y); const y = 20;");
} catch (e) {
    __check(__line("TDZ const ReferenceError"), "TDZ const ReferenceError");
}
