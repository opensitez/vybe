// vybe-test: js/object_prototype_has_own_property_is_prototype_of/test_js_object_has_own_null_prototype_safety
// origin: languages/js/tests/js/test_js_object_prototype_has_own_property_is_prototype_of.rs

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

const obj = Object.create(null);
obj.prop = 42;

// Object.hasOwn is safe on null prototype objects!
__check(__line(Object.hasOwn(obj, "prop")), "true");

try {
    // Calling hasOwnProperty directly on object with null prototype throws TypeError!
    obj.hasOwnProperty("prop");
} catch (e) {
    __check(__line("Direct Call Failed"), "Direct Call Failed");
}
