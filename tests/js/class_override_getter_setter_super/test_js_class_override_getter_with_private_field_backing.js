// vybe-test: js/class_override_getter_setter_super/test_js_class_override_getter_with_private_field_backing
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
    #val = 50;
    get val() { return this.#val; }
    set val(v) { this.#val = v; }
}
class Derived extends Base {
    get val() { return super.val * 2; }
}
const d = new Derived();
d.val = 100;
__check(__line(d.val), "200");
