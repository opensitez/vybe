// vybe-test: js/class_private_advanced/private_method_accessible_within_class_only
// origin: languages/js/tests/js/test_class_private_advanced.rs

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
    #isNonEmpty(s) { return s.length > 0; }
    validate(s) { return this.#isNonEmpty(s) ? "ok" : "empty"; }
}
const v = new Validator();
__check(__line(v.validate("hello")), "ok");
__check(__line(v.validate("")), "empty");
