// vybe-test: js/class_private_deep/test_private_static_method_in_static_getter
// origin: languages/js/tests/js/test_class_private_deep.rs

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
    static #compute() { return "staticSecret"; }
    static get secret() { return Secret.#compute(); }
}
__check(__line(Secret.secret), "staticSecret");
