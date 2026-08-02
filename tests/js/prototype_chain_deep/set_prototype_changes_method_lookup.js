// vybe-test: js/prototype_chain_deep/set_prototype_changes_method_lookup
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

const animal = { speak() { return "generic animal sound"; } };
const dog = { speak() { return "woof"; } };
const pet = {};
Object.setPrototypeOf(pet, animal);
__check(__line(pet.speak()), "generic animal sound");
Object.setPrototypeOf(pet, dog);
__check(__line(pet.speak()), "woof");
