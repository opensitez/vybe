// vybe-test: js/class_static_private_fields_methods/test_js_class_static_private_method_name_property
// origin: languages/js/tests/js/test_js_class_static_private_fields_methods.rs

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

class Metadata {
    static #internalMethod() {}
    static getName() { return Metadata.#internalMethod.name; }
}
__check(__line(Metadata.getName()), "#internalMethod");
