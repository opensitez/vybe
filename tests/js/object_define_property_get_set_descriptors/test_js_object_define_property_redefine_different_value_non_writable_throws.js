// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_property_redefine_different_value_non_writable_throws
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
Object.defineProperty(obj, "v", {
    value: 99,
    writable: false,
    configurable: false
});
try {
    Object.defineProperty(obj, "v", { value: 100 });
} catch (e) {
    __check(__line("Redefine Failed"), "Redefine Failed");
}
