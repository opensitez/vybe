// vybe-test: js/ecma_classes/class_inheritance
// origin: languages/js/tests/js/test_ecma_classes.rs

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
    constructor(color) {
        this.color = color;
    }
    describe() {
        return "A " + this.color + " shape";
    }
}
class Circle extends Shape {
    constructor(color, radius) {
        super(color);
        this.radius = radius;
    }
    area() {
        return 3.14159 * this.radius * this.radius;
    }
}
const c = new Circle("red", 5);
__check(__line(c.describe()), "A red shape");
__check(__line(c.area()), "78.53975");
