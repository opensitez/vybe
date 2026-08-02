// vybe-test: js/classes/test_class_constructor_primitive_return_ignored
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

class Prim {
            constructor() {
                this.val = 42;
                return 123;
            }
        }
        const obj = new Prim();
        __check(__line(obj.val), "42");
