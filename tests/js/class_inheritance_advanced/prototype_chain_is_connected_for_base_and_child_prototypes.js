// vybe-test: js/class_inheritance_advanced/prototype_chain_is_connected_for_base_and_child_prototypes
// origin: languages/js/tests/js/test_class_inheritance_advanced.rs

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

class Base {}
class Child extends Base {}
__check(__line(Object.getPrototypeOf(Child.prototype) === Base.prototype), "true");
__check(__line(Object.getPrototypeOf(Child) === Base), "true");
const c = new Child();
__check(__line(c instanceof Base), "true");
