// vybe-test: js/class_inheritance_deep/derived_constructor_without_super_throws_reference_error
// origin: languages/js/tests/js/test_class_inheritance_deep.rs

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

class Base {}
class Broken extends Base {
    constructor() {
        this.x = 1;
    }
}

let threw = false;
try {
    new Broken();
} catch (e) {
    threw = e instanceof ReferenceError;
}
__check(__line(threw), "true");
