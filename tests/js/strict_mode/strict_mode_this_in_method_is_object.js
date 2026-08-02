// vybe-test: js/strict_mode/strict_mode_this_in_method_is_object
// origin: languages/js/tests/js/test_strict_mode.rs

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
const obj = {
    name: "test",
    getName() { return this.name; }
};
__check(__line(obj.getName()), "test");
