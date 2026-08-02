// vybe-test: js/inheritance/test_09_derived_adds_method_base_still_works
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
            hello() { return "hi"; }
        }
        class Derived extends Base {
            constructor() { super(); }
            goodbye() { return "bye"; }
        }
        let d = new Derived();
        __check(__line(d.hello(), d.goodbye()), "hi bye");
