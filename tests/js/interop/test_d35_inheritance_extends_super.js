// vybe-test: js/interop/test_d35_inheritance_extends_super
// origin: languages/js/tests/js/js_interop_test.rs

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
            constructor(name) { this.name = name; }
            describe() { return "I am a " + this.name; }
        }
        class Circle extends Shape {
            constructor(r) {
                super("circle");
                this.radius = r;
            }
        }
        let c = new Circle(5);
        __check(__line(c.describe(), c.radius), "I am a circle 5");
