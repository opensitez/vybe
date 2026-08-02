// vybe-test: js/class_patterns/computed_getter_setter
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

class Circle {
    constructor(radius) { this._radius = radius; }
    get radius() { return this._radius; }
    set radius(r) {
        if (r < 0) throw new Error("negative");
        this._radius = r;
    }
    get area() { return 3.14 * this._radius * this._radius; }
    get circumference() { return 2 * 3.14 * this._radius; }
}
let c = new Circle(5);
__check(__line(c.area), "78.5");
__check(__line(c.circumference), "31.400000000000002");
c.radius = 10;
__check(__line(c.area), "314");
