// vybe-test: js/class_inheritance_advanced/computed_super_method_call_in_instance_method
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

class Base {
    label() {
        return "base-label";
    }
}

class Child extends Base {
    getLabel() {
        const key = "label";
        return super[key]();
    }
}

const c = new Child();
__check(__line(c.getLabel()), "base-label");
