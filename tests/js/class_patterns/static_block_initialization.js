// vybe-test: js/class_patterns/static_block_initialization
// origin: languages/js/tests/js/test_class_patterns.rs

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
    static values;
    static {
        Config.values = { debug: false, version: "1.0" };
    }
}
__check(__line(Config.values.version), "1.0");
__check(__line(Config.values.debug), "false");
