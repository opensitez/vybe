// vybe-test: js/class_extends_super_constructor_call/test_js_class_instanceof_multilevel_inheritance
// origin: languages/js/tests/js/test_js_class_extends_super_constructor_call.rs

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

class GrandParent {}
class Parent extends GrandParent {}
class Child extends Parent {}

const c = new Child();
__check(__line((c instanceof Child) + "|" + (c instanceof Parent) + "|" + (c instanceof GrandParent) + "|" + (c instanceof Object)), "true|true|true|true");
