// vybe-test: js/class_extends_super_constructor_call/test_js_class_extends_null_prototype
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

class NullBase extends null {
    constructor() {
        return Object.create(null);
    }
}
const nb = new NullBase();
__check(__line(Object.getPrototypeOf(nb) === null), "true");
