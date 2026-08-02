// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_subclass_shadowing
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

class Parent {
    #action() { return "ParentPrivate"; }
    callParent() { return this.#action(); }
}
class Child extends Parent {
    #action() { return "ChildPrivate"; }
    callChild() { return this.#action(); }
}
const c = new Child();
__check(__line(`${c.callParent()}|${c.callChild()}`), "ParentPrivate|ChildPrivate");
