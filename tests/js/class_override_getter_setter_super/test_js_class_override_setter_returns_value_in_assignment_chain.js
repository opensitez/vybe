// vybe-test: js/class_override_getter_setter_super/test_js_class_override_setter_returns_value_in_assignment_chain
// origin: languages/js/tests/js/test_js_class_override_getter_setter_super.rs

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

class Base {
    set x(v) { this._x = v; }
    get x() { return this._x; }
}
class Derived extends Base {
    set x(v) { super.x = v; }
}
const d = new Derived();
const assigned = (d.x = 99);
__check(__line(assigned + "|" + d.x), "99|99");
