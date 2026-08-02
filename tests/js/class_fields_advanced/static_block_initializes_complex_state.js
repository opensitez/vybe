// vybe-test: js/class_fields_advanced/static_block_initializes_complex_state
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
    static #data = new Map([["a", 1], ["b", 2]]);
    static get(key) { return Config.#data.get(key); }
}
__check(__line(Config.get("a")), "1");
__check(__line(Config.get("b")), "2");
