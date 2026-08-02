// vybe-test: js/object_descriptors/define_properties_throws_nonextensible_in_strict_mode
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
const obj = Object.preventExtensions({});
let threw = false;
try {
    Object.defineProperties(obj, {
        a: { value: 1 },
    });
} catch {
    threw = true;
}
__check(__line(threw), "false");
__check(__line(obj.a), "1");
