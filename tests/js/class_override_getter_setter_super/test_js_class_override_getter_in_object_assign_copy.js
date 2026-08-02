// vybe-test: js/class_override_getter_setter_super/test_js_class_override_getter_in_object_assign_copy
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
    get id() { return 123; }
}
class Derived extends Base {
    get id() { return super.id + 1; }
}
const d = new Derived();
const copy = Object.assign({}, d);
__check(__line(copy.id), "124");
