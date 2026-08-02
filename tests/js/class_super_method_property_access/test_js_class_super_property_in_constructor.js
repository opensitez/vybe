// vybe-test: js/class_super_method_property_access/test_js_class_super_property_in_constructor
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
    initMessage() { return "Initialized"; }
}
class Sub extends Base {
    constructor() {
        super();
        this.msg = super.initMessage();
    }
}
__check(__line(new Sub().msg), "Initialized");
