// vybe-test: js/object_prevent_extensions_seal_freeze/test_js_object_freeze_shallow_only
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

const obj = {
    nested: { val: 5 }
};
Object.freeze(obj);
obj.nested.val = 50; // Nested object is NOT frozen!
__check(__line(obj.nested.val), "50");
