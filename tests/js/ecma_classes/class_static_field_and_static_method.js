// vybe-test: js/ecma_classes/class_static_field_and_static_method
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
    static label = "base";
    static next() {
        return Config.label + "-next";
    }
}
__check(__line(Config.label), "base");
__check(__line(Config.next()), "base-next");
Config.label = "override";
__check(__line(Config.next()), "override-next");
