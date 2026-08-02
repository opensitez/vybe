// vybe-test: js/classes/test_class_property_set
// origin: languages/js/tests/js/js_classes_test.rs

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

class Dog {
            constructor(name) {
                this.name = name;
                this.tricks = 0;
            }
            learn() {
                this.tricks = this.tricks + 1;
            }
        }
        let d = new Dog("Rex");
        d.learn();
        d.learn();
        __check(__line(d.name, d.tricks), "Rex 2");
