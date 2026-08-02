// vybe-test: js/class_static_private_fields_methods/test_js_class_static_private_method_subclass_shadowing
// origin: languages/js/tests/js/test_js_class_static_private_fields_methods.rs

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
    static #fn() { return "BaseStatic"; }
    static callBase() { return Base.#fn(); }
}
class Sub extends Base {
    static #fn() { return "SubStatic"; }
    static callSub() { return Sub.#fn(); }
}
__check(__line(`${Base.callBase()}|${Sub.callSub()}`), "BaseStatic|SubStatic");
