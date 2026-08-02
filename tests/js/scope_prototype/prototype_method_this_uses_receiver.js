// vybe-test: js/scope_prototype/prototype_method_this_uses_receiver
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

function Person(name) { this.name = name; }
Person.prototype.greet = function() { return "hi " + this.name; };
let a = new Person("Alice");
let b = { name: "Bob", greet: a.greet };
__check(__line(a.greet()), "hi Alice");
__check(__line(b.greet()), "hi Bob");
