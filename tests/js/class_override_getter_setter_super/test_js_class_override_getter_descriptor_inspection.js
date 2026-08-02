// vybe-test: js/class_override_getter_setter_super/test_js_class_override_getter_descriptor_inspection
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
    get item() { return "B"; }
}
class Derived extends Base {
    get item() { return "D"; }
}
const desc = Object.getOwnPropertyDescriptor(Derived.prototype, "item");
__check(__line(typeof desc.get + "|" + desc.enumerable + "|" + desc.configurable), "function|false|true");
