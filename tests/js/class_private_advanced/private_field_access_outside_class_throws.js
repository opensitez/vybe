// vybe-test: js/class_private_advanced/private_field_access_outside_class_throws
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

class Secret {
    #value = 99;
    getValue() { return this.#value; }
}
const s = new Secret();
console.log(s.getValue());
try {
    console.log(s.#value);
} catch (e) {
    console.log("access denied");
}
