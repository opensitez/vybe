// vybe-test: js/prototype_oop_patterns/mixin_composition
// origin: languages/js/tests/js/test_prototype_oop_patterns.rs

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

const Serializable = (superclass) => class extends superclass {
    serialize() { return JSON.stringify(this); }
};
const Timestamped = (superclass) => class extends superclass {
    constructor(...args) { super(...args); this.createdAt = 0; }
};
class Base {
    constructor(name) { this.name = name; }
}
class User extends Timestamped(Serializable(Base)) {}
const u = new User("Alice");
const s = u.serialize();
__check(__line(JSON.parse(s).name), "Alice");
__check(__line("createdAt" in u), "true");
__check(__line(u instanceof Base), "true");
