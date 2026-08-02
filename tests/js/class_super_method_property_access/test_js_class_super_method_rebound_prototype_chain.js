// vybe-test: js/class_super_method_property_access/test_js_class_super_method_rebound_prototype_chain
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

class Base1 { action() { return "B1"; } }
class Base2 { action() { return "B2"; } }
class Sub extends Base1 {
    action() { return super.action(); }
}
const s = new Sub();
Object.setPrototypeOf(Sub.prototype, Base2.prototype);
__check(__line(s.action()), "B2"); // HomeObject[[Prototype]] is Base2 -> returns "B2"!
