// vybe-test: js/class_super_method_property_access/test_js_class_super_property_read_own_instance_field_returns_undefined
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

class Base {
    baseField = "InstanceField";
}
class Sub extends Base {
    checkSuperField() {
        return super.baseField === undefined; // super looks up on Prototype chain, NOT instance fields!
    }
}
__check(__line(new Sub().checkSuperField()), "true");
