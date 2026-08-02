// vybe-test: js/class_super_method_property_access/test_js_class_super_method_this_receiver_preservation
// origin: languages/js/tests/js/test_js_class_super_method_property_access.rs

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
    getName() { return this.name; }
}
class Child extends Parent {
    constructor(name) {
        super();
        this.name = name;
    }
    getName() {
        return super.getName().toUpperCase(); // super.getName() executes with 'this' pointing to Child instance!
    }
}
__check(__line(new Child("alice").getName()), "ALICE");
