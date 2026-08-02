// vybe-test: js/prototype_oop_patterns/parasitic_inheritance
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

function createEnhanced(original) {
    const clone = Object.create(original);
    clone.describe = function() { return "Enhanced: " + this.name; };
    return clone;
}
const base = { name: "Base", greet() { return "Hello from " + this.name; } };
const enhanced = createEnhanced(base);
enhanced.name = "Enhanced";
__check(__line(enhanced.greet()), "Hello from Enhanced");
__check(__line(enhanced.describe()), "Enhanced: Enhanced");
__check(__line(Object.getPrototypeOf(enhanced) === base), "true");
