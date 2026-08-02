// vybe-test: js/scope_prototype/prototype_method_lookup
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
Animal.prototype.speak = function() { return this.name + " speaks"; };
let a = new Animal("Dog");
__check(__line(a.speak()), "Dog speaks");
__check(__line(a.hasOwnProperty("name")), "true");
__check(__line(a.hasOwnProperty("speak")), "false");
