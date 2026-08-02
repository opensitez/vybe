// vybe-test: js/interop/test_a02_method_call_on_class_instance
// origin: languages/js/tests/js/js_interop_test.rs

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

class Greeter {
            constructor(name) { this.name = name; }
            greet() { return "Hello " + this.name; }
        }
        let g = new Greeter("World");
        __check(__line(g.greet()), "Hello World");
