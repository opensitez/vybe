// vybe-test: js/class_inheritance_advanced/test_static_super_method_call_with_modified_arguments
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

class Base {
    static greet(name) {
        return "Hello " + name;
    }
}
class Child extends Base {
    static greet(name) {
        return super.greet(name.toUpperCase());
    }
}
__check(__line(Child.greet("world")), "Hello WORLD");
