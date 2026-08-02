// vybe-test: js/class_inheritance_advanced/base_prototype_for_instance_fields_after_subclass_fields
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
    constructor() {}
}
class Child extends Base {
    childField = "child";
    constructor() {
        super();
        this.baseAndChild = this.baseField + "|" + this.childField;
    }
}
const c = new Child();
__check(__line(c.baseAndChild), "base|child");
