// vybe-test: js/class_private_advanced/private_static_method_helper_in_public_factory
// origin: languages/js/tests/js/test_class_private_advanced.rs

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

class UUID {
    static #pad(n) { return n.toString(16).padStart(4, "0"); }
    static #segment(max) { return Math.floor(max / 2); }
    static create(seed) {
        const a = UUID.#pad(seed);
        const b = UUID.#pad(UUID.#segment(seed));
        return a + "-" + b;
    }
}
__check(__line(UUID.create(256)), "0100-0080");
__check(__line(UUID.create(65536)), "10000-8000");
