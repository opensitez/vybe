// vybe-test: js/oop_patterns_advanced/abstract_class_simulation
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

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
class Circle extends Shape {
    constructor(r) { super(); this.r = r; }
    area() { return Math.PI * this.r * this.r; }
}
const c = new Circle(1);
__check(__line(Math.abs(c.area() - Math.PI) < 0.0001), "true");
let threw = false;
try { new Shape().area(); } catch { threw = true; }
__check(__line(threw), "true");
