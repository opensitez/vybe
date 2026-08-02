// vybe-test: js/oop_patterns_advanced/interface_duck_typing
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

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

function implements_(obj, methods) {
    return methods.every(m => typeof obj[m] === "function");
}
const Serializable = ["serialize", "deserialize"];
class Config {
    serialize() { return JSON.stringify(this.data); }
    deserialize(s) { this.data = JSON.parse(s); return this; }
}
const cfg = new Config();
cfg.data = { x: 1 };
__check(__line(implements_(cfg, Serializable)), "true");
__check(__line(implements_({}, Serializable)), "false");
