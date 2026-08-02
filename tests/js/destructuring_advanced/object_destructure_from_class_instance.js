// vybe-test: js/destructuring_advanced/object_destructure_from_class_instance
// origin: languages/js/tests/js/test_destructuring_advanced.rs

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
const { x, y } = new Point(3, 4);
__check(__line(x + y), "7");
