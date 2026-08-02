// vybe-test: js/inheritance/test_19_derived_overrides_getter
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
            constructor() { this._x = 5; }
            get x() { return this._x; }
        }
        class Derived extends Base {
            constructor() { super(); }
            get x() { return this._x * 2; }
        }
        let b = new Base();
        let d = new Derived();
        __check(__line(b.x, d.x), "5 10");
