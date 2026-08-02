// vybe-test: js/class_static_deep/static_private_field
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

class Registry {
    static #instances = 0;
    static create() {
        Registry.#instances++;
        return new Registry();
    }
    static getCount() { return Registry.#instances; }
}
Registry.create();
Registry.create();
__check(__line(Registry.getCount()), "2");
