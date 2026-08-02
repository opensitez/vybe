// vybe-test: js/class_inheritance_deep/three_level_inheritance_chain
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
    who() { return "A"; }
}
class B extends A {
    who() { return super.who() + "B"; }
}
class C extends B {
    who() { return super.who() + "C"; }
}
__check(__line(new C().who()), "ABC");
