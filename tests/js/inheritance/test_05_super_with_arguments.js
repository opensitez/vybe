// vybe-test: js/inheritance/test_05_super_with_arguments
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

class Shape {
            constructor(type, sides) {
                this.type = type;
                this.sides = sides;
            }
        }
        class Triangle extends Shape {
            constructor() {
                super("triangle", 3);
            }
        }
        let t = new Triangle();
        __check(__line(t.type, t.sides), "triangle 3");
