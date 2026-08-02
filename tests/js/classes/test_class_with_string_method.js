// vybe-test: js/classes/test_class_with_string_method
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

class Greeter {
            constructor(name) {
                this.name = name;
            }
            greet() {
                return "Hello, " + this.name + "!";
            }
        }
        let g = new Greeter("World");
        __check(__line(g.greet()), "Hello, World!");
