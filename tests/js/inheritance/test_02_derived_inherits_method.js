// vybe-test: js/inheritance/test_02_derived_inherits_method
// origin: languages/js/tests/js/js_inheritance_test.rs

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
            constructor(x) { this.x = x; }
            getX() { return this.x; }
        }
        class Derived extends Base {
            constructor(x) { super(x); }
        }
        let d = new Derived(42);
        __check(__line(d.getX()), "42");
