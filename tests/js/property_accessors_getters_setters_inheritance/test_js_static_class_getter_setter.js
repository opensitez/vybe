// vybe-test: js/property_accessors_getters_setters_inheritance/test_js_static_class_getter_setter
// origin: languages/js/tests/js/test_js_property_accessors_getters_setters_inheritance.rs

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
    static _env = "dev";
    static get env() { return this._env; }
    static set env(v) { this._env = v; }
}
Config.env = "prod";
__check(__line(Config.env), "prod");
