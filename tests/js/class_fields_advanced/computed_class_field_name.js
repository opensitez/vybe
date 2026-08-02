// vybe-test: js/class_fields_advanced/computed_class_field_name
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

const fieldName = "dynamic";
class Dyn {
    constructor() { this[fieldName] = 42; }
}
const d = new Dyn();
__check(__line(d.dynamic), "42");
