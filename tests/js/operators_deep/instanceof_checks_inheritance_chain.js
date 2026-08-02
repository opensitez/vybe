// vybe-test: js/operators_deep/instanceof_checks_inheritance_chain
// origin: languages/js/tests/js/test_operators_deep.rs

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
class Derived extends Base {}
const value = new Derived();
__check(__line(value instanceof Derived), "true");
__check(__line(value instanceof Base), "true");
__check(__line(42 instanceof Base), "false");
