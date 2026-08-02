// vybe-test: js/design_patterns/singleton_pattern
// origin: languages/js/tests/js/test_design_patterns.rs

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
    static #instance = null;
    #settings = {};
    static getInstance() {
        if (!Config.#instance) Config.#instance = new Config();
        return Config.#instance;
    }
    set(key, val) { this.#settings[key] = val; return this; }
    get(key) { return this.#settings[key]; }
}
const a = Config.getInstance();
const b = Config.getInstance();
a.set("theme", "dark");
__check(__line(a === b), "true");
__check(__line(b.get("theme")), "dark");
