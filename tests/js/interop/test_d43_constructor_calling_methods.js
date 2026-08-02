// vybe-test: js/interop/test_d43_constructor_calling_methods
// origin: languages/js/tests/js/js_interop_test.rs

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

class Validator {
            constructor(value) {
                this.value = value;
                this.valid = this.check();
            }
            check() { return this.value > 0; }
        }
        let v1 = new Validator(10);
        let v2 = new Validator(-5);
        __check(__line(v1.valid, v2.valid), "true false");
