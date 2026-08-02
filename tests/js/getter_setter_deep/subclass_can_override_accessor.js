// vybe-test: js/getter_setter_deep/subclass_can_override_accessor
// origin: languages/js/tests/js/test_getter_setter_deep.rs

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
    get label() { return "Base"; }
}
class Child extends Base {
    get label() { return "Child:" + super.label; }
}
__check(__line(new Child().label), "Child:Base");
