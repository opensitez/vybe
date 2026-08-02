// vybe-test: js/class_inheritance_deep/abstract_class_throws_when_instantiated
// origin: languages/js/tests/js/test_class_inheritance_deep.rs

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

class AbstractShape {
    constructor() {
        if (this.constructor === AbstractShape) {
            throw new Error("Cannot instantiate abstract class");
        }
    }
    area() { throw new Error("Must implement area()"); }
}
class Square extends AbstractShape {
    constructor(side) { super(); this.side = side; }
    area() { return this.side ** 2; }
}
let threw = false;
try { new AbstractShape(); } catch (e) { threw = true; }
__check(__line(threw), "true");
const s = new Square(4);
__check(__line(s.area()), "16");
