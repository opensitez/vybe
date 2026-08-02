// vybe-test: js/class_static_private_fields_methods/test_js_class_static_private_field_this_receiver_check
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

class Parent {
    static #secret = "ParentSecret";
    static getSecret() {
        return this.#secret; // 'this' must be Parent class constructor!
    }
}
class Child extends Parent {}

__check(__line(Parent.getSecret()), "ParentSecret");
try {
    Child.getSecret(); // Called on Child constructor where #secret does NOT exist -> throws TypeError!
} catch (e) {
    __check(__line("Child Subclass Static Private Call Error"), "Child Subclass Static Private Call Error");
}
