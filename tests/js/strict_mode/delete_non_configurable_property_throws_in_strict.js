// vybe-test: js/strict_mode/delete_non_configurable_property_throws_in_strict
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
const obj = {};
Object.defineProperty(obj, "x", { value: 1, configurable: false });
let threw = false;
try { delete obj.x; } catch (e) { threw = e instanceof TypeError; }
__check(__line(threw), "true");
