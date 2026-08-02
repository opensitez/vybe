// vybe-test: js/class_private_advanced/class_field_non_private_public
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

class Dog {
    species = "canine";
    constructor(name) { this.name = name; }
    describe() { return this.name + " is a " + this.species; }
}
const d = new Dog("Rex");
__check(__line(d.describe()), "Rex is a canine");
__check(__line(d.species), "canine");
