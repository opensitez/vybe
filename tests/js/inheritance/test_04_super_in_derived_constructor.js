// vybe-test: js/inheritance/test_04_super_in_derived_constructor
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

class Parent {
            constructor() { this.role = "parent"; }
        }
        class Kid extends Parent {
            constructor() {
                super();
                this.role2 = "kid";
            }
        }
        let k = new Kid();
        __check(__line(k.role, k.role2), "parent kid");
