// vybe-test: js/interop/test_a09_constructor_sets_properties
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

class Point {
            constructor(x, y) {
                this.x = x;
                this.y = y;
                this.label = "P(" + x + "," + y + ")";
            }
        }
        let p = new Point(3, 4);
        __check(__line(p.x, p.y, p.label), "3 4 P(3,4)");
