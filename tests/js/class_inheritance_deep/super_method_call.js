// vybe-test: js/class_inheritance_deep/super_method_call
// origin: languages/js/tests/js/test_class_inheritance_deep.rs

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
    greet() { return "Base"; }
}
class Child extends Base {
    greet() { return super.greet() + "+Child"; }
}
__check(__line(new Child().greet()), "Base+Child");
