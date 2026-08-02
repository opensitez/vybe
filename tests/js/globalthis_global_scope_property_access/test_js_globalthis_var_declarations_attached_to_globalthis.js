// vybe-test: js/globalthis_global_scope_property_access/test_js_globalthis_var_declarations_attached_to_globalthis
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

var explicitVar = "attached";
__check(__line(globalThis.explicitVar), "attached");
