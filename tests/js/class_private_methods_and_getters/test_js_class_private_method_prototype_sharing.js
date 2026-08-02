// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_prototype_sharing
// origin: languages/js/tests/js/test_js_class_private_methods_and_getters.rs

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

class Foo {
    #shared() { return 42; }
    getShared(other) { return other.#shared(); }
}
const f1 = new Foo();
const f2 = new Foo();
// Private methods are shared on brand check registry across instances
__check(__line(f1.getShared(f2)), "42");
