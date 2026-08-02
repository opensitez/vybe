// vybe-test: js/object_create_prototype_descriptors/test_js_object_create_property_descriptor_must_be_object
// origin: languages/js/tests/js/test_js_object_create_prototype_descriptors.rs

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
    Object.create(null, { prop: "not_an_object_descriptor" });
} catch (e) {
    __check(__line("Invalid Property Descriptor TypeError"), "Invalid Property Descriptor TypeError");
}
