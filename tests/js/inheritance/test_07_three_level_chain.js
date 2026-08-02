// vybe-test: js/inheritance/test_07_three_level_chain
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
            constructor() { this.a = "A"; }
            whoA() { return this.a; }
        }
        class B extends A {
            constructor() {
                super();
                this.b = "B";
            }
            whoB() { return this.b; }
        }
        class C extends B {
            constructor() {
                super();
                this.c = "C";
            }
        }
        let c = new C();
        __check(__line(c.whoA(), c.whoB(), c.c), "A B C");
