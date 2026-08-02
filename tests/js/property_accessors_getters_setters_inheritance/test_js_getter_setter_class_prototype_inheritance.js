// vybe-test: js/property_accessors_getters_setters_inheritance/test_js_getter_setter_class_prototype_inheritance
// origin: languages/js/tests/js/test_js_property_accessors_getters_setters_inheritance.rs

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
    constructor(radius) { this.radius = radius; }
    get area() { return Math.PI * this.radius ** 2; }
}
const c = new Circle(2);
__check(__line(c.area.toFixed(2)), "12.57");
