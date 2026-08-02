// vybe-test: js/class_inheritance_deep/test_static_super_method_preserves_dynamic_this_receiver
// origin: languages/js/tests/js/test_class_inheritance_deep.rs

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
    static getName() {
        return this.name;
    }
}
class Derived extends Base {
    static getName() {
        return super.getName() + "Suffix";
    }
}
__check(__line(Derived.getName()), "DerivedSuffix");
