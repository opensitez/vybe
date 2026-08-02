// vybe-test: js/class_extends_super_constructor_call/test_js_class_extends_builtin_map_subclassing
// origin: languages/js/tests/js/test_js_class_extends_super_constructor_call.rs

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

class DefaultMap extends Map {
    get(key) {
        if (!this.has(key)) this.set(key, 0);
        return super.get(key);
    }
}
const m = new DefaultMap();
__check(__line(m.get("counter")), "0");
