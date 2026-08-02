// vybe-test: js/class_inheritance_advanced/constructor_returning_object_override_skips_instance_initialization
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
    constructor() {
        this.tag = "base";
    }
}

class Child extends Base {
    constructor() {
        super();
        return { tag: "replacement", marker: "overridden" };
    }
}

const c = new Child();
__check(__line(c.tag), "replacement");
__check(__line(c.marker), "overridden");
__check(__line(c instanceof Base), "false");
