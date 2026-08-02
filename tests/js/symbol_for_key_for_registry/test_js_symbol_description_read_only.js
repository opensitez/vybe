// vybe-test: js/symbol_for_key_for_registry/test_js_symbol_description_read_only
// origin: languages/js/tests/js/test_js_symbol_for_key_for_registry.rs

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

const sym = Symbol("orig");
try {
    "use strict";
    sym.description = "new";
} catch (e) {
    __check(__line("Description Read-Only TypeError"), "Description Read-Only TypeError");
}
