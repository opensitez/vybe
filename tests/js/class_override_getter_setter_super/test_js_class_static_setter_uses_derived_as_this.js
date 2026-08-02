// vybe-test: js/class_override_getter_setter_super/test_js_class_static_setter_uses_derived_as_this
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
    static set level(v) { this._setBy = this.name; this._level = v; }
    static get level() { return `${this._setBy}|${this._level}`; }
}
class Derived extends Base {
    static set level(v) { super.level = v; }
}
Derived.level = 42;
__check(__line(Derived.level), "Derived|42");
