// vybe-test: js/ecma/test_class_extends_override
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

class Shape {
            constructor(type) {
                this.type = type;
            }
            describe() {
                return "I am a " + this.type;
            }
        }
        class Circle extends Shape {
            constructor(radius) {
                super();
                this.type = "circle";
                this.radius = radius;
            }
            area() {
                return Math.PI * this.radius * this.radius;
            }
        }
        let c = new Circle(5);
        __check(__line(c.type), "circle");
        __check(__line(Math.round(c.area())), "79");
