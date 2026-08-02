// vybe-test: js/class_patterns/abstract_class_pattern
// origin: languages/js/tests/js/test_class_patterns.rs

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
    area() { throw new Error("not implemented"); }
}
class Rect extends Shape {
    constructor(w, h) { super(); this.w = w; this.h = h; }
    area() { return this.w * this.h; }
}
class Circle extends Shape {
    constructor(r) { super(); this.r = r; }
    area() { return 3.14 * this.r * this.r; }
}
let shapes = [new Rect(3, 4), new Circle(5)];
shapes.forEach(s => console.log(s.area()));
try {
    new Shape().area();
} catch (e) {
    console.log(e.message);
}
