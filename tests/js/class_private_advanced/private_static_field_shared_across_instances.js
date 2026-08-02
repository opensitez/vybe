// vybe-test: js/class_private_advanced/private_static_field_shared_across_instances
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

class Registry {
    static #count = 0;
    constructor() { Registry.#count++; }
    static getCount() { return Registry.#count; }
}
new Registry();
new Registry();
new Registry();
__check(__line(Registry.getCount()), "3");
