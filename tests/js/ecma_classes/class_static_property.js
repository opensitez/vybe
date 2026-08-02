// vybe-test: js/ecma_classes/class_static_property
// origin: languages/js/tests/js/test_ecma_classes.rs

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
    static version = "1.0.0";
    static debug = false;
}
__check(__line(Config.version), "1.0.0");
__check(__line(Config.debug), "false");
