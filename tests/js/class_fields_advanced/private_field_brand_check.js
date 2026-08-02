// vybe-test: js/class_fields_advanced/private_field_brand_check
// origin: languages/js/tests/js/test_class_fields_advanced.rs

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
    #radius;
    constructor(r) { this.#radius = r; }
    static isCircle(obj) { return #radius in obj; }
    area() { return Math.PI * this.#radius ** 2; }
}
const c = new Circle(5);
__check(__line(Circle.isCircle(c)), "true");
__check(__line(Circle.isCircle({})), "false");
