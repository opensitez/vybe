// vybe-test: js/class_inheritance_advanced/super_getter_uses_base_implementation
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
    get label() {
        return "base";
    }
}

class Child extends Base {
    constructor(tag) {
        super();
        this.tag = tag;
    }

    get label() {
        return super.label + "|" + this.tag;
    }
}

const c = new Child("node");
__check(__line(c.label), "base|node");
__check(__line(c instanceof Base), "true");
