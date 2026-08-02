// vybe-test: js/strict_mode_invariants_delete_arguments_this/test_js_strict_mode_assignment_to_read_only_property_throws_typeerror
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

const obj = {};
Object.defineProperty(obj, "fixed", { value: 10, writable: false });
try {
    "use strict";
    obj.fixed = 20;
} catch (e) {
    __check(__line("Strict Read-Only Property TypeError"), "Strict Read-Only Property TypeError");
}
