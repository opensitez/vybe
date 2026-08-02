// vybe-test: js/delete_operator/delete_non_configurable_property_throws_typeerror_in_strict_mode
// origin: languages/js/tests/js/test_delete_operator.rs

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

'use strict';
const obj = {};
Object.defineProperty(obj, 'locked', {
    value: 1,
    configurable: false,
});
try {
    delete obj.locked;
} catch (e) {
    __check(__line(e.name), "TypeError");
}
