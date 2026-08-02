// vybe-test: js/mixin_abstract_patterns/super_method_call_in_mixin
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

const Logger = Base => class extends Base {
    greet() {
        return "[LOG] " + super.greet();
    }
};
class Person {
    constructor(name) { this.name = name; }
    greet() { return "Hi, I'm " + this.name; }
}
class LoggedPerson extends Logger(Person) {}
const p = new LoggedPerson("Bob");
__check(__line(p.greet()), "[LOG] Hi, I'm Bob");
