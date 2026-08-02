// vybe-test: js/inheritance/test_20_no_constructor_auto_super
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
            constructor() { this.ready = true; }
        }
        class Derived extends Base {}
        let d = new Derived();
        __check(__line(d.ready), "true");
