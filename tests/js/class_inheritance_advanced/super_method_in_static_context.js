// vybe-test: js/class_inheritance_advanced/super_method_in_static_context
// origin: languages/js/tests/js/test_class_inheritance_advanced.rs

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

class Animal {
    static describe() { return "Animal"; }
}
class Dog extends Animal {
    static describe() { return super.describe() + "/Dog"; }
}
__check(__line(Dog.describe()), "Animal/Dog");
