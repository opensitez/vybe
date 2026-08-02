// vybe-test: js/class_private_in_operator_brand_check/test_js_private_in_operator_inherited_prototype_returns_false
// origin: languages/js/tests/js/test_js_class_private_in_operator_brand_check.rs

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
    #parentField = 100;
    static hasParentField(obj) {
        return #parentField in obj;
    }
}
class Child extends Parent {}

const p = new Parent();
const c = new Child();
__check(__line(Parent.hasParentField(p) + "|" + Parent.hasParentField(c)), "true|true");
