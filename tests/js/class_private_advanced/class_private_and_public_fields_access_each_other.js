// vybe-test: js/class_private_advanced/class_private_and_public_fields_access_each_other
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

class User {
    name;
    #role = "user";
    constructor(name, role) {
        this.name = name;
        if (role) this.#role = role;
    }
    display() { return this.name + " [" + this.#role + "]"; }
}
const u1 = new User("Alice", "admin");
const u2 = new User("Bob");
__check(__line(u1.display()), "Alice [admin]");
__check(__line(u2.display()), "Bob [user]");
__check(__line(u1.name), "Alice");
