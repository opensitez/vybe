// vybe-test: js/ecma_classes/class_multi_level_inheritance
// origin: languages/js/tests/js/test_ecma_classes.rs

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
    whoami() { return "A"; }
}
class B extends A {
    whoami() { return "B->" + super.whoami(); }
}
class C extends B {
    whoami() { return "C->" + super.whoami(); }
}
const c = new C();
__check(__line(c.whoami()), "C->B->A");
