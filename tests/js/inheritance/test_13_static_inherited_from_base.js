// vybe-test: js/inheritance/test_13_static_inherited_from_base
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

class Base {
            static helper() { return "from base"; }
        }
        class Derived extends Base {
            constructor() { super(); }
        }
        __check(__line(Derived.helper()), "from base");
