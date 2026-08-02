// vybe-test: js/ecma_classes/class_getter
// origin: languages/js/tests/js/test_ecma_classes.rs

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

class Rectangle {
    constructor(w, h) {
        this.width = w;
        this.height = h;
    }
    get area() {
        return this.width * this.height;
    }
}
const r = new Rectangle(5, 3);
__check(__line(r.area), "15");
