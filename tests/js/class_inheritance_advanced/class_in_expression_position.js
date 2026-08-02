// vybe-test: js/class_inheritance_advanced/class_in_expression_position
// origin: languages/js/tests/js/test_class_inheritance_advanced.rs

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

const Greeter = class NamedGreeter {
    greet(name) { return "Hello " + name; }
};
const g = new Greeter();
__check(__line(g.greet("World")), "Hello World");
__check(__line(typeof g.greet), "function");
