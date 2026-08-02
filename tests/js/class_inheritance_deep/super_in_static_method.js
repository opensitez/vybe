// vybe-test: js/class_inheritance_deep/super_in_static_method
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

class A {
    static who() { return "A"; }
}
class B extends A {
    static who() { return super.who() + "B"; }
}
__check(__line(B.who()), "AB");
