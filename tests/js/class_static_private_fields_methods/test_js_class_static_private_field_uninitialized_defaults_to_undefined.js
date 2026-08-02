// vybe-test: js/class_static_private_fields_methods/test_js_class_static_private_field_uninitialized_defaults_to_undefined
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

class Holder {
    static #data;
    static check() { return Holder.#data === undefined; }
}
__check(__line(Holder.check()), "true");
