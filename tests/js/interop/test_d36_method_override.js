// vybe-test: js/interop/test_d36_method_override
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

class Base {
            greet() { return "base"; }
        }
        class Child extends Base {
            greet() { return "child"; }
        }
        let b = new Base();
        let c = new Child();
        __check(__line(b.greet(), c.greet()), "base child");
