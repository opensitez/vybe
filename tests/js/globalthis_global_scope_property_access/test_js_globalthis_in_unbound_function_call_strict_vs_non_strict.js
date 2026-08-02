// vybe-test: js/globalthis_global_scope_property_access/test_js_globalthis_in_unbound_function_call_strict_vs_non_strict
// origin: languages/js/tests/js/test_js_globalthis_global_scope_property_access.rs

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

function nonStrictThis() { return this; }
function strictThis() { "use strict"; return this; }

__check(__line((nonStrictThis() === globalThis) + "|" + (strictThis() === undefined)), "true|true");
