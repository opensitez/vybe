// vybe-test: js/dynamic_import_promise_resolution/test_js_dynamic_import_invalid_specifier_rejection
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
        await import("invalid_non_existent_module_specifier_xyz_123");
    } catch (e) {
        console.log("ImportRejected");
    }
})();
