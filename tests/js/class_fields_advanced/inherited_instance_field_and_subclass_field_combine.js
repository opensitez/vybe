// vybe-test: js/class_fields_advanced/inherited_instance_field_and_subclass_field_combine
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

class Base {
    base = "base";
}

class Child extends Base {
    child = "child";
}

const c = new Child();
__check(__line(c.base), "base");
__check(__line(c.child), "child");
