// vybe-test: js/class_override_getter_setter_super/test_js_class_setter_override_without_super_hides_parent_setter
// origin: languages/js/tests/js/test_js_class_override_getter_setter_super.rs

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
    set data(v) { this.parentData = v; }
}
class Child extends Parent {
    set data(v) { this.childData = v; }
}
const c = new Child();
c.data = "Test";
__check(__line(c.parentData + "|" + c.childData), "undefined|Test");
