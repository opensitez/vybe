// vybe-test: js/class_private_advanced/instanceof_with_private_fields
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

class Animal {
    #alive = true;
    isAlive() { return this.#alive; }
}
class Dog extends Animal {
    #breed;
    constructor(breed) { super(); this.#breed = breed; }
    getBreed() { return this.#breed; }
}
const d = new Dog("Labrador");
__check(__line(d instanceof Dog), "true");
__check(__line(d instanceof Animal), "true");
__check(__line(d.isAlive()), "true");
__check(__line(d.getBreed()), "Labrador");
