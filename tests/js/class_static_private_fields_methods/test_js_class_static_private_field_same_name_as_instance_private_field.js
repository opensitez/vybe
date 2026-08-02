// vybe-test: js/class_static_private_fields_methods/test_js_class_static_private_field_same_name_as_instance_private_field
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

class Hybrid {
    #val = "InstancePrivate";
    static #val = "StaticPrivate";

    getInstanceVal() { return this.#val; }
    static getStaticVal() { return Hybrid.#val; }
}
const h = new Hybrid();
__check(__line(`${h.getInstanceVal()}|${Hybrid.getStaticVal()}`), "InstancePrivate|StaticPrivate");
