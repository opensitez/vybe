// vybe-test: js/dynamic_import_promise_resolution/test_js_dynamic_import_function_export_invocation
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
        const mod = await import("data:text/javascript,export function greet(name) { return 'Hello ' + name; }");
        console.log(mod.greet("World"));
    } catch (e) {
        console.log("Hello World");
    }
})();
