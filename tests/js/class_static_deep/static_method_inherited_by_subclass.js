// vybe-test: js/class_static_deep/static_method_inherited_by_subclass
// origin: languages/js/tests/js/test_class_static_deep.rs

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
    static describe() { return "I am " + this.name; }
}
class Dog extends Animal {}
__check(__line(Dog.describe()), "I am Dog");
