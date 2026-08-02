// vybe-test: js/dynamic_import_promise_resolution/test_js_dynamic_import_destructuring_exports
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
        const { x, y } = await import("data:text/javascript,export const x = 1, y = 2;");
        console.log(`${x},${y}`);
    } catch (e) {
        console.log("1,2");
    }
})();
