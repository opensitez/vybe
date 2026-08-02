// vybe-test: js/class_fields_advanced/test_derived_field_initializers_run_after_super_returns
// origin: languages/js/tests/js/test_class_fields_advanced.rs

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

let baseSeen = "uninitialized";
class Base {
    constructor() {
        baseSeen = String(this.derivedField);
    }
}
class Derived extends Base {
    derivedField = "initialized";
}
const d = new Derived();
__check(__line(`${baseSeen}|${d.derivedField}`), "undefined|initialized");
