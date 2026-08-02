// vybe-test: js/object_descriptors/define_property_non_configurable_prevents_redefinition
// origin: languages/js/tests/js/test_object_descriptors.rs

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
Object.defineProperty(obj, "locked", { value: 1, configurable: false });
let threw = false;
try {
    Object.defineProperty(obj, "locked", { value: 2 });
} catch { threw = true; }
__check(__line(threw), "true");
