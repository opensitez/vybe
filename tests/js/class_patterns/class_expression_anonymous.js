// vybe-test: js/class_patterns/class_expression_anonymous
// origin: languages/js/tests/js/test_class_patterns.rs

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

let Animal = class {
    constructor(name) { this.name = name; }
    speak() { return this.name + " speaks"; }
};
let a = new Animal("Cat");
__check(__line(a.speak()), "Cat speaks");
