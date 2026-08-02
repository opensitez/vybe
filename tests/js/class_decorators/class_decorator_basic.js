// vybe-test: js/class_decorators/class_decorator_basic
// origin: languages/js/tests/js/test_class_decorators.rs

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

function sealed(target) { Object.seal(target.prototype); return target; }
const Point = sealed(class {
    constructor(x, y) { this.x = x; this.y = y; }
});
const p = new Point(1, 2);
__check(__line(p.x), "1");
__check(__line(Object.isSealed(Point.prototype)), "true");
