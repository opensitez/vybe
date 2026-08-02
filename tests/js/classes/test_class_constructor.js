// vybe-test: js/classes/test_class_constructor
// origin: languages/js/tests/js/js_classes_test.rs

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
        }
        let p = new Point(3, 4);
        __check(__line(p.x, p.y), "3 4");
