// vybe-test: js/classes/test_class_inheritance_super_call
// origin: languages/js/tests/js/js_classes_test.rs

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
            constructor(name) {
                this.kind = "base";
                this.name = name;
            }
            describe() {
                return this.kind + ":" + this.name;
            }
        }
        class Derived extends Base {
            constructor(name, role) {
                super(name);
                this.role = role;
            }
            describe() {
                return super.describe() + ":" + this.role;
            }
        }
        const d = new Derived("alpha", "admin");
        __check(__line(d.describe()), "base:alpha:admin");
