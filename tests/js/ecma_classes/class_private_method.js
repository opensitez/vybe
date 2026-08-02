// vybe-test: js/ecma_classes/class_private_method
// origin: languages/js/tests/js/test_ecma_classes.rs

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
    #validate(input) {
        return input.length > 0;
    }
    check(input) {
        return this.#validate(input);
    }
}
const v = new Validator();
__check(__line(v.check("hello")), "true");
__check(__line(v.check("")), "false");
