// vybe-test: js/class_private_advanced/static_init_block_can_call_static_methods
// origin: languages/js/tests/js/test_class_private_advanced.rs

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
    static value;
    static #compute() { return 42; }
    static {
        Config.value = Config.#compute();
    }
}
__check(__line(Config.value), "42");
