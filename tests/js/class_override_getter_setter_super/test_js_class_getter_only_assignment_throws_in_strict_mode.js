// vybe-test: js/class_override_getter_setter_super/test_js_class_getter_only_assignment_throws_in_strict_mode
// origin: languages/js/tests/js/test_js_class_override_getter_setter_super.rs

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

"use strict";
class Base {
    get val() { return 10; }
}
class Derived extends Base {}
const d = new Derived();
try {
    d.val = 20;
} catch (e) {
    __check(__line(e.name), "TypeError");
}
