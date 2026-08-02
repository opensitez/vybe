// vybe-test: js/class_super_method_property_access/test_js_class_super_property_increment_decrement_operators
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
    get count() { return this._c || 0; }
    set count(v) { this._c = v; }
}
class Sub extends Base {
    increment() {
        super.count++; // Performs super.count = super.count + 1 on 'this'
    }
}
const s = new Sub();
s.increment();
s.increment();
__check(__line(s.count), "2");
