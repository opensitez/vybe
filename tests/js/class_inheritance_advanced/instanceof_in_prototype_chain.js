// vybe-test: js/class_inheritance_advanced/instanceof_in_prototype_chain
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

class A {}
class B extends A {}
class C extends B {}
const c = new C();
__check(__line(c instanceof C), "true");
__check(__line(c instanceof B), "true");
__check(__line(c instanceof A), "true");
__check(__line(c instanceof Object), "true");
const b = new B();
__check(__line(b instanceof C), "false"); // false — b is not a C
