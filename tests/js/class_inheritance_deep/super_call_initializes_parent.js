// vybe-test: js/class_inheritance_deep/super_call_initializes_parent
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
    constructor(name) { this.name = name; }
    speak() { return this.name + " speaks"; }
}
class Dog extends Animal {
    constructor(name, breed) {
        super(name);
        this.breed = breed;
    }
}
const d = new Dog("Rex", "Lab");
__check(__line(d.name), "Rex");
__check(__line(d.breed), "Lab");
__check(__line(d.speak()), "Rex speaks");
