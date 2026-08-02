// vybe-test: js/class_fields_advanced/static_field_shared_on_class
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
    static defaultTimeout = 5000;
    static VERSION = "1.0.0";
}
__check(__line(Config.defaultTimeout), "5000");
__check(__line(Config.VERSION), "1.0.0");
// Not on instances
const c = new Config();
__check(__line(c.defaultTimeout === undefined), "true");
