// vybe-test: js/inheritance/test_10_instanceof_through_chain
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

class A {}
        class B extends A {
            constructor() { super(); }
        }
        class C extends B {
            constructor() { super(); }
        }
        let c = new C();
        __check(__line(c instanceof C, c instanceof B, c instanceof A), "true true true");
