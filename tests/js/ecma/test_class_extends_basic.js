// vybe-test: js/ecma/test_class_extends_basic
// origin: languages/js/tests/js/js_ecma_test.rs

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
                return this.name + " makes a sound";
            }
        }
        class Dog extends Animal {
            constructor(name) {
                super();
                this.name = name;
            }
            bark() {
                return this.name + " barks";
            }
        }
        let d = new Dog("Rex");
        __check(__line(d.name), "Rex");
        __check(__line(d.bark()), "Rex barks");
