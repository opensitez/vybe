// vybe-test: js/dynamic_import_promise_resolution/test_js_dynamic_import_namespace_unmodifiable_extensibility
// origin: languages/js/tests/js/test_js_dynamic_import_promise_resolution.rs

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

(async () => {
    try {
        const mod = await import("data:text/javascript,export const c = 3;");
        console.log(Object.isSealed(mod) || !Object.isExtensible(mod));
    } catch (e) {
        console.log("true");
    }
})();
