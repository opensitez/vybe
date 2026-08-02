// vybe-test: js/object_immutability/prevent_extensions_freeze_status_checks
// origin: languages/js/tests/js/test_object_immutability.rs

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

const obj = { a: 1, b: 2 };
Object.preventExtensions(obj);
__check(__line(Object.isExtensible(obj)), "false");
__check(__line(Object.isSealed(obj)), "false");
__check(__line(Object.isFrozen(obj)), "false");
Object.freeze(obj);
__check(__line(Object.isExtensible(obj)), "false");
__check(__line(Object.isSealed(obj)), "true");
__check(__line(Object.isFrozen(obj)), "true");
