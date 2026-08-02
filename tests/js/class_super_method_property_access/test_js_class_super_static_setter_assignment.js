// vybe-test: js/class_super_method_property_access/test_js_class_super_static_setter_assignment
// origin: languages/js/tests/js/test_js_class_super_method_property_access.rs

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
    static set marker(v) {
        this._marker = `base:${v}`;
    }
    static get marker() {
        return this._marker;
    }
}
class Derived extends Base {
    static applyMarker(v) {
        super.marker = v;
    }
}

Derived.applyMarker("X");
__check(__line(Derived.marker), "base:X");
__check(__line(Base.marker), "undefined");
__check(__line(Object.hasOwn(Derived, "_marker")), "true");
__check(__line(Object.hasOwn(Base, "_marker")), "false");
