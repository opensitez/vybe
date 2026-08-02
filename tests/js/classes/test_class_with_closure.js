// vybe-test: js/classes/test_class_with_closure
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

class Adder {
            constructor(base) {
                this.base = base;
            }
            add(x) {
                return this.base + x;
            }
        }
        let a = new Adder(100);
        __check(__line(a.add(42)), "142");
