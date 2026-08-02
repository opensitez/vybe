// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_instanceof_uses_current_prototype_property_of_rhs
// origin: languages/js/tests/js/test_js_relational_in_instanceof_less_greater_operators.rs

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

function C() {}
const before = new C();
__check(__line(before instanceof C), "true");

C.prototype = {};
const after = new C();
__check(__line(after instanceof C), "true");
__check(__line(before instanceof C), "false");
