// vybe-test: js/class_super_method_property_access/test_js_class_super_method_apply_call_receiver
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
    greet() { return `Hello ${this.name}`; }
}
class Sub extends Base {
    greet() {
        return super.greet.call({ name: "CustomReceiver" });
    }
}
__check(__line(new Sub().greet()), "Hello CustomReceiver");
