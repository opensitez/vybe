// vybe-test: js/inheritance/test_08_constructor_chain_three_levels
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

class A {
            constructor(v) { this.val = v; }
        }
        class B extends A {
            constructor(v) {
                super(v * 2);
            }
        }
        class C extends B {
            constructor(v) {
                super(v + 1);
            }
        }
        let c = new C(3);
        __check(__line(c.val), "8");
