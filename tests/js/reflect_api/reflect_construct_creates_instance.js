// vybe-test: js/reflect_api/reflect_construct_creates_instance
// origin: languages/js/tests/js/test_reflect_api.rs

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

class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
const p = Reflect.construct(Point, [3, 4]);
__check(__line(p.x), "3");
__check(__line(p.y), "4");
__check(__line(p instanceof Point), "true");
