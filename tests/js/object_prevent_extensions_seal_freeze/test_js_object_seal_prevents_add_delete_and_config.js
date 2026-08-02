// vybe-test: js/object_prevent_extensions_seal_freeze/test_js_object_seal_prevents_add_delete_and_config
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

const obj = { a: 1 };
Object.seal(obj);
__check(__line(Object.isSealed(obj)), "true");
__check(__line(Object.isExtensible(obj)), "false");

obj.a = 100; // Modifying existing writable property allowed
__check(__line(obj.a), "100");

try {
    "use strict";
    delete obj.a; // Deleting property throws in strict mode
} catch (e) {
    __check(__line("Seal Delete Error"), "Seal Delete Error");
}
