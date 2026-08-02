// vybe-test: js/inheritance/test_derived_constructor_missing_super_call_throws_referenceerror
// origin: languages/js/tests/js/js_inheritance_test.rs

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
        class Derived extends Base {
            constructor() {}
        }
        try {
            new Derived();
        } catch (e) {
            __check(__line(e.name), "ReferenceError");
        }
