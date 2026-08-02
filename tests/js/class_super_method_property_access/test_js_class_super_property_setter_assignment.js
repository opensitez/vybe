// vybe-test: js/class_super_method_property_access/test_js_class_super_property_setter_assignment
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
    set data(v) { this._data = v + 10; }
    get data() { return this._data; }
}
class Derived extends Base {
    setData(v) {
        super.data = v; // super.data = v sets property on 'this' receiver using Base's setter!
    }
}
const d = new Derived();
d.setData(50);
console.log(d.data);
