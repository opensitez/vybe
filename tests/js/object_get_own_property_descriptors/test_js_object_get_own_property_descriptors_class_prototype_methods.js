// vybe-test: js/object_get_own_property_descriptors/test_js_object_get_own_property_descriptors_class_prototype_methods
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

Class Foo {
    bar() {}
}
const descs = Object.getOwnPropertyDescriptors(Foo.prototype);
__check(__line(descs.bar.enumerable + "|" + descs.bar.configurable), "false|true");
