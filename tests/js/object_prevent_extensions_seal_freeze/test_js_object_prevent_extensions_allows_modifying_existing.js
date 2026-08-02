// vybe-test: js/object_prevent_extensions_seal_freeze/test_js_object_prevent_extensions_allows_modifying_existing
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

const obj = { val: 10 };
Object.preventExtensions(obj);
obj.val = 20;
__check(__line(obj.val), "20");
delete obj.val;
__check(__line(obj.val), "undefined");
