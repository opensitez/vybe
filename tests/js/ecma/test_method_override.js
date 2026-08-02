// vybe-test: js/ecma/test_method_override
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
            constructor() { this.type = "shape"; }
            describe() { return "I am a " + this.type; }
        }
        class Circle extends Shape {
            constructor(r) {
                super();
                this.type = "circle";
                this.r = r;
            }
            describe() { return "Circle with radius " + this.r; }
        }
        let s = new Shape();
        let c = new Circle(5);
        __check(__line(s.describe()), "I am a shape");
        __check(__line(c.describe()), "Circle with radius 5");
