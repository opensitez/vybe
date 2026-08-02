// vybe-test: js/prototype_patterns_deep/prototype_assignment_via_object_create
// origin: languages/js/tests/js/test_prototype_patterns_deep.rs

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

const animal = {
    speak() { return this.name + " says " + this.sound; }
};
const dog = Object.create(animal);
dog.name = "Rex";
dog.sound = "woof";
__check(__line(dog.speak()), "Rex says woof");
