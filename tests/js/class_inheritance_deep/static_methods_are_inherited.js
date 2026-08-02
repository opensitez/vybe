// vybe-test: js/class_inheritance_deep/static_methods_are_inherited
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
    type() { return "base"; }
}
class Child extends Base {
    type() { return "child"; }
}
Child.create = function() { return new Child(); };
const obj = Child.create();
__check(__line(obj instanceof Child), "true");
__check(__line(obj.type()), "child");
