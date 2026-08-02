// vybe-test: js/class_override_getter_setter_super/test_js_class_override_setter_only
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
    set val(v) { this._val = v + 1; }
    get val() { return this._val; }
}
class Derived extends Base {
    set val(v) { super.val = v * 10; }
}
const d = new Derived();
d.val = 5;
__check(__line(d.val), "51");
