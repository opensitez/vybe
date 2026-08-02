// vybe-test: js/class_override_getter_setter_super/test_js_class_override_getter_validation_throws
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
    get age() { return this._age; }
    set age(v) { this._age = v; }
}
class Derived extends Base {
    set age(v) {
        if (v < 0) throw new RangeError("Negative Age Error");
        super.age = v;
    }
}
const d = new Derived();
d.age = 20;
__check(__line(d.age), "20");
try {
    d.age = -5;
} catch (e) {
    __check(__line("RangeError Caught"), "RangeError Caught");
}
