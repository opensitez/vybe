// vybe-test: js/strict_mode/delete_configurable_property_in_strict_mode_works
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
const obj = { a: 1 };
const result = delete obj.a;
__check(__line(result), "true");
__check(__line("a" in obj), "false");
