// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_property_value_writable_false
// origin: languages/js/tests/js/test_js_object_define_property_get_set_descriptors.rs

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
Object.defineProperty(obj, "prop", {
    value: 42,
    writable: false,
    configurable: true,
    enumerable: true
});
__check(__line(obj.prop), "42");
try {
    "use strict";
    obj.prop = 99;
} catch (e) {
    __check(__line("Error: " + e.name), "Error: TypeError");
}
__check(__line(obj.prop), "42");
