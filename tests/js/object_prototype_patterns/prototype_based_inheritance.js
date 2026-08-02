// vybe-test: js/object_prototype_patterns/prototype_based_inheritance
// origin: languages/js/tests/js/test_object_prototype_patterns.rs

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

const shape = {
    area() { return 0; },
    perimeter() { return 0; },
    describe() { return `${this.constructor.name}: area=${this.area()}`; }
};
const circle = Object.create(shape);
circle.constructor = { name: "Circle" };
circle.init = function(r) { this.r = r; return this; };
circle.area = function() { return Math.PI * this.r * this.r; };
const c = Object.create(circle).init(3);
__check(__line(c.area().toFixed(4)), "28.2743");
