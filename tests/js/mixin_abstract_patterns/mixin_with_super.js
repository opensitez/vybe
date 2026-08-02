// vybe-test: js/mixin_abstract_patterns/mixin_with_super
// origin: languages/js/tests/js/test_mixin_abstract_patterns.rs

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

const Timestamped = Base => class extends Base {
    constructor(...args) {
        super(...args);
        this.createdAt = 0;
    }
};
class User {
    constructor(name) { this.name = name; }
}
class TimestampedUser extends Timestamped(User) {}
const u = new TimestampedUser("Alice");
__check(__line(u.name), "Alice");
__check(__line(u.createdAt), "0");
