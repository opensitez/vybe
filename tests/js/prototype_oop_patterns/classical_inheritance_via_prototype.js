// vybe-test: js/prototype_oop_patterns/classical_inheritance_via_prototype
// origin: languages/js/tests/js/test_prototype_oop_patterns.rs

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
Animal.prototype.speak = function() { return this.name + " speaks"; };
function Dog(name) { Animal.call(this, name); }
Dog.prototype = Object.create(Animal.prototype);
Dog.prototype.constructor = Dog;
Dog.prototype.bark = function() { return this.name + " barks"; };
const d = new Dog("Rex");
__check(__line(d.speak()), "Rex speaks");
__check(__line(d.bark()), "Rex barks");
__check(__line(d instanceof Dog), "true");
__check(__line(d instanceof Animal), "true");
