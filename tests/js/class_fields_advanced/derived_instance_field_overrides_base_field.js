// vybe-test: js/class_fields_advanced/derived_instance_field_overrides_base_field
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

class Base {
    x = 1;
}

class Derived extends Base {
    x = 2;
}

const d = new Derived();
__check(__line(d.x), "2");
