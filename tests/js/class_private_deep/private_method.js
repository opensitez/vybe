// vybe-test: js/class_private_deep/private_method
// origin: languages/js/tests/js/test_class_private_deep.rs

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
    #validate(x) { return x > 0; }
    check(x) { return this.#validate(x) ? "valid" : "invalid"; }
}
const v = new Validator();
__check(__line(v.check(5)), "valid");
__check(__line(v.check(-1)), "invalid");
