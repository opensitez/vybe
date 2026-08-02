// vybe-test: js/class_inheritance_deep/mixin_chain_multiple
// origin: languages/js/tests/js/test_class_inheritance_deep.rs

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

const Serializable = (Base) => class extends Base {
    serialize() { return JSON.stringify(this); }
};
const Loggable = (Base) => class extends Base {
    log(msg) { return `[LOG] ${msg}`; }
};

class Model { constructor(data) { Object.assign(this, data); } }
class User extends Serializable(Loggable(Model)) {}

const u = new User({ name: "Alice", age: 30 });
__check(__line(u.log("hello")), "[LOG] hello");
const data = JSON.parse(u.serialize());
__check(__line(data.name), "Alice");
