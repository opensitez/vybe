// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_cross_instance_call
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

class EncryptedString {
    #data;
    constructor(d) { this.#data = d; }
    #getRaw() { return this.#data; }

    compare(other) {
        return this.#getRaw() === other.#getRaw();
    }
}
const s1 = new EncryptedString("ABC");
const s2 = new EncryptedString("ABC");
__check(__line(s1.compare(s2)), "true");
