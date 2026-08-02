// vybe-test: js/comprehensive/test_class_inheritance_super
// origin: languages/js/tests/js/js_comprehensive_test.rs

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
            speak() { return this.name + " makes a noise"; }
        }
        class Dog extends Animal {
            constructor(name) { super(name); }
            speak() { return this.name + " barks"; }
        }
        let d = new Dog("Rex");
        __check(__line(d.speak()), "Rex barks");
