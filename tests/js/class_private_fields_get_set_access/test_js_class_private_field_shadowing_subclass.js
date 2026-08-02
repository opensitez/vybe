// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_shadowing_subclass
// origin: languages/js/tests/js/test_js_class_private_fields_get_set_access.rs

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

class Base {
    #value = "BasePrivate";
    getBaseValue() { return this.#value; }
}
class Derived extends Base {
    #value = "DerivedPrivate"; // Distinct private field declaration!
    getDerivedValue() { return this.#value; }
}
const d = new Derived();
__check(__line(`${d.getBaseValue()}|${d.getDerivedValue()}`), "BasePrivate|DerivedPrivate");
