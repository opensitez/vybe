// vybe-test: js/class_fields_advanced/private_field_not_accessible_outside
// origin: languages/js/tests/js/test_class_fields_advanced.rs

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
    #value = 42;
    get() { return this.#value; }
}
const s = new Secret();
__check(__line(s.get()), "42");
let threw = false;
try { s.#value; } catch { threw = true; }
__check(__line(threw), "true");
