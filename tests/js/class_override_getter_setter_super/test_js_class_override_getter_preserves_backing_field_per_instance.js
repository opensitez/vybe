// vybe-test: js/class_override_getter_setter_super/test_js_class_override_getter_preserves_backing_field_per_instance
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
    constructor(v) { this._v = v; }
    get value() { return this._v; }
}
class Derived extends Base {
    get value() { return `[${super.value}]`; }
}
const d1 = new Derived("A");
const d2 = new Derived("B");
__check(__line(`${d1.value}:${d2.value}`), "[A]:[B]");
