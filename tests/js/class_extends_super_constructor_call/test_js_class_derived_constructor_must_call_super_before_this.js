// vybe-test: js/class_extends_super_constructor_call/test_js_class_derived_constructor_must_call_super_before_this
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
        try {
            eval("this.x = 10; super();");
        } catch (e) {
            __check(__line("This Access Before Super ReferenceError"), "This Access Before Super ReferenceError");
        }
    }
}
new Sub();
