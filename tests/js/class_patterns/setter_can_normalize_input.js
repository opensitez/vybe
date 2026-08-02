// vybe-test: js/class_patterns/setter_can_normalize_input
// origin: languages/js/tests/js/test_class_patterns.rs

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

class User {
    set name(value) { this._name = value.trim(); }
    get name() { return this._name; }
}
let u = new User();
u.name = "  Alice  ";
__check(__line(u.name), "Alice");
