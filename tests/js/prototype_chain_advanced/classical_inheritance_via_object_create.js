// vybe-test: js/prototype_chain_advanced/classical_inheritance_via_object_create
// origin: languages/js/tests/js/test_prototype_chain_advanced.rs

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

function Animal(name) { this.name = name; }
Animal.prototype.speak = function() { return this.name + " says hi"; };

function Dog(name, breed) {
    Animal.call(this, name);
    this.breed = breed;
}
Dog.prototype = Object.create(Animal.prototype);
Dog.prototype.constructor = Dog;
Dog.prototype.bark = function() { return this.name + " barks"; };

const d = new Dog("Rex", "Lab");
__check(__line(d.speak()), "Rex says hi");
__check(__line(d.bark()), "Rex barks");
__check(__line(d instanceof Dog), "true");
__check(__line(d instanceof Animal), "true");
