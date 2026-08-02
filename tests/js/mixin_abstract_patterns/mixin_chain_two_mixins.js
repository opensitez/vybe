// vybe-test: js/mixin_abstract_patterns/mixin_chain_two_mixins
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

const Swimmable = Base => class extends Base {
    swim() { return this.name + " swims"; }
};
const Flyable = Base => class extends Base {
    fly() { return this.name + " flies"; }
};
class Animal {
    constructor(name) { this.name = name; }
}
class Duck extends Swimmable(Flyable(Animal)) {}
const d = new Duck("Donald");
__check(__line(d.swim()), "Donald swims");
__check(__line(d.fly()), "Donald flies");
