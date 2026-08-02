// vybe-test: js/prototype_chain_deep/instanceof_checks_prototype_chain
// origin: languages/js/tests/js/test_prototype_chain_deep.rs

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

class Animal {}
class Dog extends Animal {}
const d = new Dog();
__check(__line(d instanceof Dog), "true");
__check(__line(d instanceof Animal), "true");
__check(__line(d instanceof Object), "true");
