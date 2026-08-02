// vybe-test: js/class_override_getter_setter_super/test_js_class_override_getter_returning_object
// origin: languages/js/tests/js/test_js_class_override_getter_setter_super.rs

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

class Base {
    get config() { return { port: 80 }; }
}
class Derived extends Base {
    get config() {
        const baseCfg = super.config;
        return { ...baseCfg, ssl: true };
    }
}
const d = new Derived();
__check(__line(`${d.config.port}:${d.config.ssl}`), "80:true");
