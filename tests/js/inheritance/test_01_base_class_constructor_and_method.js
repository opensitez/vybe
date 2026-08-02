// vybe-test: js/inheritance/test_01_base_class_constructor_and_method
// origin: languages/js/tests/js/js_inheritance_test.rs

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
                return "I am " + this.name;
            }
        }
        let a = new Animal("dog");
        __check(__line(a.speak()), "I am dog");
