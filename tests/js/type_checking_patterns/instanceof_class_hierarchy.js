// vybe-test: js/type_checking_patterns/instanceof_class_hierarchy
// origin: languages/js/tests/js/test_type_checking_patterns.rs

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
const obj = new C();
__check(__line(obj instanceof C), "true");
__check(__line(obj instanceof B), "true");
__check(__line(obj instanceof A), "true");
__check(__line(obj instanceof Object), "true");
