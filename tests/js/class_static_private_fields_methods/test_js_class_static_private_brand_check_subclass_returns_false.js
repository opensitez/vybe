// vybe-test: js/class_static_private_fields_methods/test_js_class_static_private_brand_check_subclass_returns_false
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
    static #brand = true;
    static isParent(target) {
        return #brand in target;
    }
}
class Child extends Parent {}
__check(__line(Parent.isParent(Parent) + "|" + Parent.isParent(Child)), "true|false");
