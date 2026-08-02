// vybe-test: js/class_inheritance_advanced/test_super_property_assignment_targets_receiver_this
// origin: languages/js/tests/js/test_class_inheritance_advanced.rs

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

class Base {}
class Child extends Base {
    setProp(v) {
        super.x = v;
    }
}
const c = new Child();
c.setProp(42);
__check(__line(c.x + "|" + ("x" in Base.prototype)), "42|false");
