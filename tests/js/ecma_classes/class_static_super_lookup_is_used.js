// vybe-test: js/ecma_classes/class_static_super_lookup_is_used
// origin: languages/js/tests/js/test_ecma_classes.rs

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
    static label() { return "base"; }
}
class Derived extends Base {
    static label() { return super.label() + "/derived"; }
}
__check(__line(Derived.label()), "base/derived");
__check(__line(Object.getPrototypeOf(Derived) === Base), "true");
