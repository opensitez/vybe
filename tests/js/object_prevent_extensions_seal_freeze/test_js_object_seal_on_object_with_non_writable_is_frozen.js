// vybe-test: js/object_prevent_extensions_seal_freeze/test_js_object_seal_on_object_with_non_writable_is_frozen
// origin: languages/js/tests/js/test_js_object_prevent_extensions_seal_freeze.rs

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

const obj = {};
Object.defineProperty(obj, "prop", {
    value: 10,
    writable: false,
    configurable: true
});
Object.seal(obj);
__check(__line(Object.isFrozen(obj)), "true");
