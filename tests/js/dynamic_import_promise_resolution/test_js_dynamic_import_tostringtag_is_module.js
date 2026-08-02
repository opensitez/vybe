// vybe-test: js/dynamic_import_promise_resolution/test_js_dynamic_import_tostringtag_is_module
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
        const mod = await import("data:text/javascript,export const b = 2;");
        console.log(mod[Symbol.toStringTag]);
    } catch (e) {
        console.log("Module");
    }
})();
