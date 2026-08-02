// vybe-test: js/class_fields_advanced/class_static_fields_do_not_inherit_instance_shape
// origin: languages/js/tests/js/test_class_fields_advanced.rs

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
    static role = "parent";
    instanceRole = "instance-parent";
}

class Child extends Parent {
    static role = "child";
}

const p = new Parent();
const c = new Child();
__check(__line(p.instanceRole), "instance-parent");
__check(__line(Parent.role), "parent");
__check(__line(Child.role), "child");
__check(__line(c.instanceRole), "instance-parent");
