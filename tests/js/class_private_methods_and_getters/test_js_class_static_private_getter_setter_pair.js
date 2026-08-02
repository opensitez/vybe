// vybe-test: js/class_private_methods_and_getters/test_js_class_static_private_getter_setter_pair
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

class System {
    static #val = 0;
    static get #secret() { return System.#val; }
    static set #secret(v) { System.#val = v; }

    static update(v) {
        System.#secret = v;
        return System.#secret;
    }
}
__check(__line(System.update(42)), "42");
