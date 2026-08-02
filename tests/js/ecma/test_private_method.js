// vybe-test: js/ecma/test_private_method
// origin: languages/js/tests/js/js_ecma_test.rs

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
            #isValid(value) {
                return value > 0;
            }
            validate(value) {
                if (this.#isValid(value)) {
                    return "valid";
                }
                return "invalid";
            }
        }
        let v = new Validator();
        __check(__line(v.validate(5), v.validate(-1)), "valid invalid");
