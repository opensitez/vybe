// vybe-test: js/class_extends_super_constructor_call/test_js_class_super_called_twice_throws_referenceerror
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
        try {
            super(); // Double super call throws ReferenceError!
        } catch (e) {
            __check(__line("Double Super ReferenceError"), "Double Super ReferenceError");
        }
    }
}
new Sub();
