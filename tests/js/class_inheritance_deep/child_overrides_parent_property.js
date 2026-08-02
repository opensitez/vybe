// vybe-test: js/class_inheritance_deep/child_overrides_parent_property
// origin: languages/js/tests/js/test_class_inheritance_deep.rs

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
    get name() { return "Shape"; }
    area() { return 0; }
}
class Circle extends Shape {
    constructor(r) { super(); this.r = r; }
    get name() { return "Circle"; }
    area() { return Math.PI * this.r * this.r; }
}
const c = new Circle(1);
__check(__line(c.name), "Circle");
__check(__line(c.area().toFixed(5)), "3.14159");
