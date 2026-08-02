// vybe-test: js/class_patterns/mixin_pattern
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

let Serializable = (Base) => class extends Base {
    serialize() { return JSON.stringify(this); }
};
let Loggable = (Base) => class extends Base {
    log() { __check(__line("LOG: " + this.name), "LOG: Alice"); }
};
class User {
    constructor(name) { this.name = name; }
}
class EnhancedUser extends Loggable(Serializable(User)) {}
let u = new EnhancedUser("Alice");
u.log();
let s = u.serialize();
__check(__line(s.includes("Alice")), "true");
