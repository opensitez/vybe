// vybe-test: js/property_accessors_getters_setters_inheritance/test_js_super_getter_call_in_derived_class
// origin: languages/js/tests/js/test_js_property_accessors_getters_setters_inheritance.rs

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
    get label() { return "ParentLabel"; }
}
class Child extends Parent {
    get label() { return super.label + " -> ChildLabel"; }
}
const c = new Child();
__check(__line(c.label), "ParentLabel -> ChildLabel");
