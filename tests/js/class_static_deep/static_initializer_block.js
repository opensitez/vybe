// vybe-test: js/class_static_deep/static_initializer_block
// origin: languages/js/tests/js/test_class_static_deep.rs

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
        Config.values = [1, 2, 3];
        Config.sum = Config.values.reduce((a, b) => a + b, 0);
    }
}
__check(__line(Config.sum), "6");
__check(__line(Config.values.join(",")), "1,2,3");
