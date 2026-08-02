// vybe-test: js/class_static_deep/static_init_block_can_access_private
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

class Secret {
    static #value = 42;
    static get() { return Secret.#value; }
}
__check(__line(Secret.get()), "42");
