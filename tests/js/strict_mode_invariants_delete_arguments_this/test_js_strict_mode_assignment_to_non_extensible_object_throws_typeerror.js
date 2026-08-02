// vybe-test: js/strict_mode_invariants_delete_arguments_this/test_js_strict_mode_assignment_to_non_extensible_object_throws_typeerror
// origin: languages/js/tests/js/test_js_strict_mode_invariants_delete_arguments_this.rs

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

const obj = Object.preventExtensions({});
try {
    "use strict";
    obj.newProp = 10;
} catch (e) {
    __check(__line("Strict Non-Extensible Assignment TypeError"), "Strict Non-Extensible Assignment TypeError");
}
