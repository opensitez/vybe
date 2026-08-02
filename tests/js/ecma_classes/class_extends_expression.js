// vybe-test: js/ecma_classes/class_extends_expression
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

function getBase() {
    return class {
        greet() { return "base"; }
    };
}
class Derived extends getBase() {
    greet() { return super.greet() + "+derived"; }
}
const d = new Derived();
__check(__line(d.greet()), "base+derived");
