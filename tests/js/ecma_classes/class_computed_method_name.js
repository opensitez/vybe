// vybe-test: js/ecma_classes/class_computed_method_name
// origin: languages/js/tests/js/test_ecma_classes.rs

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

const methodName = "greet";
class Greeter {
    [methodName]() {
        return "Hello!";
    }
}
const g = new Greeter();
__check(__line(g.greet()), "Hello!");
