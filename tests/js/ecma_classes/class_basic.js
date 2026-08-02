// vybe-test: js/ecma_classes/class_basic
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

class Animal {
    constructor(name) {
        this.name = name;
    }
    speak() {
        return this.name + " makes a noise";
    }
}
const a = new Animal("Dog");
__check(__line(a.speak()), "Dog makes a noise");
