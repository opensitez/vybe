// vybe-test: js/globalthis_global_scope_property_access/test_js_globalthis_defines_global_variable
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

globalThis.customGlobalVar = 12345;
__check(__line(customGlobalVar + "|" + globalThis.customGlobalVar), "12345|12345");
