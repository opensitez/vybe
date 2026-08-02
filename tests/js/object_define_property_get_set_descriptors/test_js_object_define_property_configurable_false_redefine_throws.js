// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_property_configurable_false_redefine_throws
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
Object.defineProperty(obj, "fixed", {
    value: 100,
    configurable: false
});
try {
    Object.defineProperty(obj, "fixed", { configurable: true });
} catch (e) {
    __check(__line("Cannot reconfigure: " + e.name), "Cannot reconfigure: TypeError");
}
