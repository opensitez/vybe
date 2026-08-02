// vybe-test: js/class_super_method_property_access/test_js_class_super_method_bound_function
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
    compute(x) { return x + 10; }
}
class Sub extends Base {
    getBoundCompute() {
        return super.compute.bind(this);
    }
}
const s = new Sub();
const bound = s.getBoundCompute();
__check(__line(bound(5)), "15");
