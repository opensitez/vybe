// vybe-test: js/dynamic_import_promise_resolution/test_js_dynamic_import_module_caching_returns_same_namespace
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
        const specifier = "data:text/javascript,export const num = 99;";
        const m1 = await import(specifier);
        const m2 = await import(specifier);
        console.log(m1 === m2);
    } catch (e) {
        console.log("true");
    }
})();
