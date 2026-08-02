// vybe-test: js/class_fields_advanced/class_public_field_is_own_enumerable_data_property
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

class Config {
    enabled = true;
}
const c = new Config();
const desc = Object.getOwnPropertyDescriptor(c, "enabled");
__check(__line(desc !== undefined), "true");
__check(__line(desc.enumerable), "true");
__check(__line(desc.writable), "true");
