// vybe-test: js/class_inheritance_deep/static_properties_inherited_by_subclass
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

class Animal {
    static count = 0;
    static increment() { Animal.count++; }
}
class Dog extends Animal {}
Dog.increment();
__check(__line(Animal.count), "1");
__check(__line(typeof Dog.count === "number"), "true");
