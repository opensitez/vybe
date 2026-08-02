// vybe-test: js/inheritance/test_17_getter_on_base
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

class Circle {
            constructor(r) { this.r = r; }
            get area() { return 3.14 * this.r * this.r; }
        }
        let c = new Circle(10);
        __check(__line(c.area), "314");
