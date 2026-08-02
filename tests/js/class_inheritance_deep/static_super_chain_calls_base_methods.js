// vybe-test: js/class_inheritance_deep/static_super_chain_calls_base_methods
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

class Animal {
    static who() { return "Animal"; }
}
class Mammal extends Animal {
    static who() { return super.who() + ":Mammal"; }
}
class Dog extends Mammal {
    static who() { return super.who() + ":Dog"; }
}
__check(__line(Dog.who()), "Animal:Mammal:Dog");
