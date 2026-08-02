// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_property_on_primitive_throws_typeerror
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

try {
    Object.defineProperty(42, "prop", { value: 1 });
} catch (e) {
    __check(__line("TypeError on Primitive"), "TypeError on Primitive");
}
