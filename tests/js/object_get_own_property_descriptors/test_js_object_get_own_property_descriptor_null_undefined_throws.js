// vybe-test: js/object_get_own_property_descriptors/test_js_object_get_own_property_descriptor_null_undefined_throws
// origin: languages/js/tests/js/test_js_object_get_own_property_descriptors.rs

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

try {
    Object.getOwnPropertyDescriptor(null, "key");
} catch (e) {
    __check(__line("Null Error"), "Null Error");
}
try {
    Object.getOwnPropertyDescriptor(undefined, "key");
} catch (e) {
    __check(__line("Undefined Error"), "Undefined Error");
}
