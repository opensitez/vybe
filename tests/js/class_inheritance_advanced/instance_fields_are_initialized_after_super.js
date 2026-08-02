// vybe-test: js/class_inheritance_advanced/instance_fields_are_initialized_after_super
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
    baseField = "base";
    constructor() { this.baseFlag = true; }
}
class Child extends Base {
    childField = "child";
    constructor() {
        super();
        this.childFlag = true;
    }
}
const c = new Child();
__check(__line(c.baseField + "|" + c.childField + "|" + c.baseFlag + "|" + c.childFlag), "base|child|true|true");
