// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_property_inheritance_accessor_receiver_this
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

const parent = {};
Object.defineProperty(parent, "name", {
    get() { return this._name || "Default"; },
    set(v) { this._name = v.toUpperCase(); },
    configurable: true
});
const child = Object.create(parent);
child.name = "alice";
__check(__line(child.name + "|" + parent.name), "ALICE|Default");
