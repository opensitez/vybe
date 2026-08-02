// vybe-test: js/inheritance/test_23_factory_static_create
// origin: languages/js/tests/js/js_inheritance_test.rs

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
            constructor(x, y) {
                this.x = x;
                this.y = y;
            }
            static origin() { return new Point(0, 0); }
        }
        let p = Point.origin();
        __check(__line(p.x, p.y), "0 0");
