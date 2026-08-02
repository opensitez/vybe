// vybe-test: js/class_private_advanced/private_getter_property
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

class Circle {
    #radius;
    constructor(r) { this.#radius = r; }
    get #area() { return 3.14 * this.#radius * this.#radius; }
    describe() { return "area=" + this.#area; }
}
const c = new Circle(5);
__check(__line(c.describe()), "area=78.5");
