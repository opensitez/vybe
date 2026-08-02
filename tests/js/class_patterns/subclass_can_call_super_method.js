// vybe-test: js/class_patterns/subclass_can_call_super_method
// origin: languages/js/tests/js/test_class_patterns.rs

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
    speak() { return "animal"; }
}
class Dog extends Animal {
    speak() { return super.speak() + " dog"; }
}
__check(__line(new Dog().speak()), "animal dog");
