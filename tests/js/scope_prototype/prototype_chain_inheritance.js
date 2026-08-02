// vybe-test: js/scope_prototype/prototype_chain_inheritance
// origin: languages/js/tests/js/test_scope_prototype.rs

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
Animal.prototype.speak = function() { return "..."; };
function Dog(name) { Animal.call(this, name); }
Dog.prototype = Object.create(Animal.prototype);
Dog.prototype.constructor = Dog;
Dog.prototype.speak = function() { return "Woof!"; };
let d = new Dog("Rex");
__check(__line(d.speak()), "Woof!");
__check(__line(d.name), "Rex");
__check(__line(d instanceof Dog), "true");
__check(__line(d instanceof Animal), "true");
