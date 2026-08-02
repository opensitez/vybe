// vybe-test: js/class_inheritance_advanced/constructor_access_to_this_before_super_throws
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
        this.base = true;
    }
}
class Child extends Base {
    constructor() {
        try {
            this.beforeSuper = true;
        } catch (e) {
            __check(__line(e.name), "ReferenceError");
            return;
        }
        super();
    }
}
new Child();
