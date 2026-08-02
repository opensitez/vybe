// vybe-test: js/class_inheritance_advanced/instance_field_initializer_uses_base_getter_via_super
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
    get name() {
        return "base";
    }
}

class Child extends Base {
    label = super.name;
    readLabel() { return super.name; }
}

const c = new Child();
__check(__line(c.label), "base");
__check(__line(c.readLabel()), "base");
