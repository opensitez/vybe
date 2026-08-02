// vybe-test: js/class_super_method_property_access/test_js_class_super_method_call_with_arguments_object
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
    add(a, b) { return a + b; }
}
class Sub extends Base {
    add() {
        return super.add(arguments[0], arguments[1]) * 10;
    }
}
__check(__line(new Sub().add(2, 3)), "50");
