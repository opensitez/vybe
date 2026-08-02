// vybe-test: js/globalthis_global_scope_property_access/test_js_globalthis_shadowing_with_local_variable
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

const globalThisLocal = "shadowed";
__check(__line(globalThisLocal + "|" + (typeof globalThis.Object !== "undefined")), "shadowed|true");
