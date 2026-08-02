// vybe-test: js/class_private_advanced/private_field_in_tostring_method
// origin: languages/js/tests/js/test_class_private_advanced.rs

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
    #x;
    #y;
    constructor(x, y) { this.#x = x; this.#y = y; }
    toString() { return "(" + this.#x + "," + this.#y + ")"; }
}
const p = new Point(3, 7);
__check(__line(p.toString()), "(3,7)");
__check(__line("Point: " + p), "Point: (3,7)");
