// vybe-test: js/object_prevent_extensions_seal_freeze/test_js_object_prevent_extensions_on_empty_object_is_sealed_and_frozen
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
Object.preventExtensions(obj);
__check(__line(Object.isExtensible(obj)), "false");
__check(__line(Object.isSealed(obj)), "true");
__check(__line(Object.isFrozen(obj)), "true");
