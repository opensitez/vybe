// vybe-test: js/object_advanced/get_prototype_of
// origin: languages/js/tests/js/test_object_advanced.rs

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

class Animal {}
class Dog extends Animal {}
let d = new Dog();
__check(__line(Object.getPrototypeOf(d) === Dog.prototype), "true");
__check(__line(Object.getPrototypeOf(Dog.prototype) === Animal.prototype), "true");
