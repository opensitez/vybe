// vybe-test: js/object_prevent_extensions_seal_freeze/test_js_object_freeze_array_mutations_throw
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

const arr = [1, 2, 3];
Object.freeze(arr);
__check(__line(Object.isFrozen(arr)), "true");
try {
    "use strict";
    arr[0] = 99;
} catch (e) {
    __check(__line("Frozen Array Element Error"), "Frozen Array Element Error");
}
try {
    arr.push(4);
} catch (e) {
    __check(__line("Frozen Array Push Error"), "Frozen Array Push Error");
}
