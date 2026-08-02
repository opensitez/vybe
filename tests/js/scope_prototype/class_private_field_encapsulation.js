// vybe-test: js/scope_prototype/class_private_field_encapsulation
// origin: languages/js/tests/js/test_scope_prototype.rs

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
    #value;
    constructor(v) { this.#value = v; }
    reveal() { return this.#value; }
}
let s = new Secret(42);
__check(__line(s.reveal()), "42");
__check(__line(s.value), "undefined");
