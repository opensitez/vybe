// vybe-test: js/globalthis_global_scope_property_access/test_js_globalthis_reflect_ownkeys_includes_standard_builtins
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

const keys = Reflect.ownKeys(globalThis);
__check(__line(keys.includes("Object") && keys.includes("Array") && keys.includes("Math")), "true");
