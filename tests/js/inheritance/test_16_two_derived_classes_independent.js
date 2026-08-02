// vybe-test: js/inheritance/test_16_two_derived_classes_independent
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
            constructor(v) { this.v = v; }
            get() { return this.v; }
        }
        class D1 extends Base {
            constructor(v) { super(v); }
        }
        class D2 extends Base {
            constructor(v) { super(v); }
        }
        let a = new D1(10);
        let b = new D2(20);
        __check(__line(a.get(), b.get()), "10 20");
