// vybe-test: js/class_extends_super_constructor_call/test_js_class_derived_ctor_returning_primitive_returns_this
// origin: languages/js/tests/js/test_js_class_extends_super_constructor_call.rs

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

class Base {}
class Sub extends Base {
    constructor() {
        super();
        return 42; // Primitive return in derived constructor is ignored!
    }
}
const s = new Sub();
__check(__line(s instanceof Sub), "true");
