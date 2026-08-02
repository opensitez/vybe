// vybe-test: js/mixin_abstract_patterns/simple_mixin_adds_methods
// origin: languages/js/tests/js/test_mixin_abstract_patterns.rs

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

const Flyable = (Base) => class extends Base {
    fly() { return this.name + " is flying"; }
};
class Animal {
    constructor(name) { this.name = name; }
}
class Bird extends Flyable(Animal) {}
const b = new Bird("Eagle");
__check(__line(b.fly()), "Eagle is flying");
__check(__line(b instanceof Animal), "true");
