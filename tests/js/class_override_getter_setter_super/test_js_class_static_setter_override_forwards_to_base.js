// vybe-test: js/class_override_getter_setter_super/test_js_class_static_setter_override_forwards_to_base
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
    static _value = 0;
    static get value() { return this._value; }
    static set value(v) { this._value = v; }
}
class Derived extends Base {
    static set value(v) { super.value = v + 7; }
}
    Derived.value = 10;
    __check(__line(`${Derived.value}`), "17");
