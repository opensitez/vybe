// vybe-test: js/class_private_methods_and_getters/test_js_class_private_setter_only_returns_undefined_on_get
// origin: languages/js/tests/js/test_js_class_private_methods_and_getters.rs

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

class WriteOnly {
    set #secret(v) { __check(__line("Set Secret: " + v), "Set Secret: Pass"); }
    setSecret(v) { this.#secret = v; }
    getSecret() {
        try {
            return this.#secret;
        } catch (e) {
            return "Get Non-Existent Getter Error";
        }
    }
}
const w = new WriteOnly();
w.setSecret("Pass");
__check(__line(w.getSecret()), "Get Non-Existent Getter Error");
